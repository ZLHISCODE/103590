Attribute VB_Name = "mdlPublic"
Option Explicit
'**************************
'       OEM����
'
'ҽҵ  D2BDD2B5
'����  CDD0C6D5
'**************************
Public Type CustomPar
    ��ʽ As Byte
    ֵ�б� As String
    ����SQL As String
    ��ϸSQL As String
    �����ֶ� As String
    ��ϸ�ֶ� As String
    ���� As String
End Type
Public gfrmMain As Object
Public gblnOK As Boolean, gblnModi As Boolean

'���ݿ���ض���
Public gcnOracle As ADODB.Connection
Public gstrDBUser As String '�û���
Public gblnDBA As Boolean '�Ƿ�DBA�û�
Public gstrUserName As String '�û�����
Public gstrUserNO As String '�û����
Public grsObject As ADODB.Recordset '��ǰ�û�������SelectȨ�޵Ķ���
'������־������ر���
Private lngErrNum As Long, strErrInfo As String, bytErrType As Byte

'API���
Public glngOldProc As Long, glngSelProc As Long
Public glngMinW As Long, glngMaxW As Long, glngMinH As Long, glngMaxH As Long
Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Const HKEY_CURRENT_USER = &H80000001
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

'����TAB���ĺ���
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE = 3

Public glngKeyHook As Long
Public gobjTab As clsTabInput
'Html Help
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Const HH_DISPLAY_TOPIC = &H0

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

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional StrName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If StrName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(StrName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(StrName)) / 2
            picBox.Print StrName
        End If
    End With
End Sub

Public Function CustomHook(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'˵����
'   Code=Hook Code(HC_ACTION��HC_NOREMOVE)
'   wParam=Virtual-Key Code
'   lParam=0-15λ(�������ظ�����)
'          16-23λ(OEM Scan Code)
'          24λ(�Ƿ���չ��,��Fx,С���̼�)
'          25-28λ(����)
'          29(ALT�Ƿ���)
'          30(������Ϣ֮ǰ���Ƿ���)
'          31(0-���ڰ���,1-�����ɿ�)
    Static blnShift As Boolean
    
    If wParam = vbKeyShift Then
        If lParam > 0 Then
            blnShift = True
        ElseIf lParam < 0 Then
            blnShift = False
        End If
    End If
    If wParam = vbKeyTab Then
        CustomHook = 1
        If blnShift Then
            If lParam > 0 Then
                gobjTab.ACT_sTabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_sTabKeyUp
            End If
        Else
            If lParam > 0 Then
                gobjTab.ACT_TabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_TabKeyUp
            End If
        End If
    Else
        CallNextHookEx glngKeyHook, Code, wParam, lParam
    End If
End Function

Public Sub RegFuncFile()
'���ܣ�ע�����������ļ�
    Dim strSys As String * 255
    
    GetSystemDirectory strSys, 255
    
    RegSetValue HKEY_CLASSES_ROOT, ".zlf", REG_SZ, "zlFunction", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction", REG_SZ, "���������ļ�", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\DefaultIcon", REG_SZ, Left(strSys, InStr(strSys, Chr(0)) - 1) & "\zl9Function.dll,0", 24
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell", REG_SZ, "Read", 4
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell\Read", REG_SZ, "�����������ļ�(&1)", 12
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell\Read\Command", REG_SZ, "NotePad.exe ""%1""", 22
End Sub

Public Function CustomMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW
        MinMax.ptMinTrackSize.Y = glngMinH
        MinMax.ptMaxTrackSize.X = glngMaxW
        MinMax.ptMaxTrackSize.Y = glngMaxH
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        CustomMessage = 1
        Exit Function
    End If
    CustomMessage = CallWindowProc(glngOldProc, hwnd, msg, wp, lp)
End Function

Public Function SelMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = 400
        MinMax.ptMinTrackSize.Y = 300
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SelMessage = 1
        Exit Function
    End If
    SelMessage = CallWindowProc(glngSelProc, hwnd, msg, wp, lp)
End Function

Public Sub ShowPercent(sngPercent As Single, objPanel As Object)
'����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ�������(��)
    Dim intAll As Integer
    intAll = objPanel.Width / frmAbout.TextWidth("��") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
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
    Else
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    End If
End Sub

Public Function CheckLen(txt As Object, intLen As Integer, strInfo As String, Optional AllowNULL As Boolean = True) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If txt.Text = "" And Not AllowNULL Then
        MsgBox "������" & strInfo & "��", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox "[" & strInfo & "]�ĳ��Ȳ��ܴ��� " & intLen & " ��", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function TLen(str As String) As Integer
'���ܣ������ַ�������ʵ����
    TLen = LenB(StrConv(str, vbFromUnicode))
End Function

Public Function TrimChar(str As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    i = InStr(strTmp, "  ")
    Do While i > 0
        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
        i = InStr(strTmp, "  ")
    Loop
    
    i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While i > 0
        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
        i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Public Sub CopyPars(ByVal objSPars As FuncPars, ByRef objOPars As FuncPars)
'���ܣ���������������
    Dim objPar As FuncPar
    
    Set objOPars = New FuncPars
    For Each objPar In objSPars
        With objPar
            objOPars.Add .����, .���, .����, .������, .����, .ȱʡֵ, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ�, .��ϸ�ֶ�, .����, "_" & .Key, .Reserve
        End With
    Next
End Sub

Public Function GetCboIndex(cbo As ComboBox, strFind As String) As Long
'���ܣ������δ�����ComboBox������ֵ
'������cbo=ComboBox,strFind=�����ַ���
    Dim i As Integer
    If strFind = "" Then GetCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strFind Then
            GetCboIndex = i
            Exit Function
        End If
    Next
    GetCboIndex = -1
End Function

Public Function CheckSQL(ByVal strSQL As String, strErr As String) As String
'���ܣ�����SQL�����д�Ƿ���ȷ
'���أ�
'     �ɹ�=SQL���ֶδ�,�����˸����ֶε����Ƽ�����,��ʽ��"����,111|����,111|����,123",����ֵ��ADO.Field.TypeΪ׼
'     ʧ��=��
    Dim rsTmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, i As Integer
    
    strCheck = strSQL
    
    If InStr(UCase(strCheck), "WHERE") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next

    Set rsTmp = zlDatabase.OpenSQLRecord(strCheck, "���SQL")
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rsTmp.Fields
            If InStr(tmpFld.Name, "|") > 0 Then
                strErr = "�ֶ�""" & tmpFld.Name & """û�б�����"
                CheckSQL = "": Exit Function
            Else
                If InStr(CheckSQL & "|", "|" & tmpFld.Name & "," & tmpFld.Type & "|") = 0 Then
                    CheckSQL = CheckSQL & "|" & tmpFld.Name & "," & tmpFld.Type
                Else
                    strErr = "������Դ�з�����ͬ���ֶ���Ŀ��"
                    CheckSQL = "": Exit Function
                End If
            End If
        Next
        CheckSQL = Mid(CheckSQL, 2)
    Else
        strErr = Err.Number & ":" & vbCrLf & Err.Description
        Err.Clear
    End If
End Function

Public Function AdjustStr(str As String) As String
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����
'˵�����Զ�(����)�����߼�"'"�綨����

    Dim i As Long, strTmp As String
    
    If InStr(1, str, "'") = 0 Then AdjustStr = "'" & str & "'": Exit Function
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(str, i, 1)
            ElseIf i = Len(str) Then
                strTmp = strTmp & Mid(str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(str, i, 1)
            End If
        End If
    Next
    AdjustStr = strTmp
End Function

Public Function MakeFile(strID As String, Optional strFormat As String = "CUSTOM") As String
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".AVI"
    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    MakeFile = strR
End Function

Public Function Currentdate() As Date
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "SELECT SYSDATE FROM DUAL"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰʱ��")
    Currentdate = rsTmp.Fields(0).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowAbout(Optional frmParent As Object)
    Dim frmShow As New frmAbout
    If frmParent Is Nothing Then
        frmShow.Show 1
    Else
        Load frmShow
        Err.Clear
        On Error Resume Next
        frmShow.Show 1, frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmShow.Show 1
        End If
    End If
End Sub

Public Function UserObject() As ADODB.Recordset
'���ܣ���ȡ��ǰ�û������в�ѯȨ�޵����б���ͼ��������(�����û�������󼰱���Ȩ����)
'���أ��ɹ�=���������б�(����Ӣ˳������),ʧ��=��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    '������."����ͼ������"
    strSQL = _
        "Select Upper(USER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW','FUNCTION')" & _
        " Union" & _
        " Select Upper(OWNER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege In('SELECT','EXECUTE')) G" & _
        " Where O.Object_Type in('TABLE','VIEW','FUNCTION')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME" & _
        " Order by OWNER,OBJECT_TYPE,OBJECT_NAME"
    
    'ALL_Object����ͼ,ֻ������ǰ�û���Ȩ�޷��ʵĶ���
    strSQL = _
        "Select Upper(USER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW','FUNCTION')" & _
        " Union" & _
        " Select Upper(OWNER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From All_Objects" & _
        " Where Object_Type in('TABLE','VIEW','FUNCTION')" & _
        " Order by OWNER,OBJECT_TYPE,OBJECT_NAME"
 
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "UserObject")
    Set UserObject = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TrueObject(ByVal strObject As String) As String
'���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim i As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    'Ѱ�Һ����һ���������ַ�
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Public Function SQLObject(ByVal strSQL As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errH
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����(��ʱstrAnalΪ�򵥲�ѯ,���ܴ�Union������)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & ",", "," & strTrue & ",") = 0 And strTrue <> "Ƕ�ײ�ѯ" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Public Function CheckObjectPriv(strObject As String, strOwner As String) As String
'���ܣ���鵱ǰ�û���ָ�������Ƿ���ȫ��Ȩ�޷���
'������strObject=��������,��"���ű�,���˷��ü�¼"
'      strOwner=������ݵ�������
'���أ���ȫ=��,����ȫ=���ܷ��ʵĶ�����,��"���ű�,���˷��ü�¼"
'˵����������У������Դ֮ǰ����Ƿ���Ȩ�޲�ѯSQL����еĶ���
'�ο���grsObject
    Dim i As Integer
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") = 0 Then
                If gblnDBA Then
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "'"
                Else
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "' And OWNER='" & UCase(strOwner) & "'"
                End If
            Else
                '�������ͼ���������ǰ׺,����������߶���Ȩ��
'                If gblnDBA Then
'                    grsObject.Filter = "OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
'                        " And OBJECT_TYPE<>'FUNCTION'"
'                Else
'                    grsObject.Filter = "OWNER='" & UCase(Split(Split(strObject, ",")(i), ".")(0)) & _
'                        "' And OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
'                        " And OBJECT_TYPE<>'FUNCTION'"
'                End If
                grsObject.Filter = "OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
                    " And OBJECT_TYPE<>'FUNCTION'"
            End If
            If grsObject.EOF Then
                If InStr(CheckObjectPriv & ",", "," & Split(strObject, ",")(i) & ",") = 0 Then
                    CheckObjectPriv = CheckObjectPriv & "," & Split(strObject, ",")(i)
                End If
            End If
        End If
    Next
    If CheckObjectPriv <> "" Then CheckObjectPriv = Mid(CheckObjectPriv, 2)
End Function

Public Function ObjectOwner(strObject As String, strOwner As String, Optional frmParent As Object) As String
'���ܣ����ݶ��������ϵ�ǰ�û����ܷ��ʵ�������ǰ׺(������ͬһ�������ж��������Ҫ��ѡ����֮һ)
'������strObject=��������,��"���ű�,���˷��ü�¼"
'���أ�����=����������ǰ׺�Ķ���,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼",ȡ��="ȡ��"
'�ο���grsObject
    Dim i As Integer, j As Integer
    
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") > 0 Then
                '�������ͼ���������ǰ׺,��ʹ���䱾����
                If InStr(ObjectOwner, "," & Split(strObject, ",")(i)) = 0 Then
                    ObjectOwner = ObjectOwner & "," & Split(strObject, ",")(i)
                End If
            Else
                If gblnDBA Then
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "'"
                Else
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "' And OWNER='" & UCase(strOwner) & "'"
                End If
                If grsObject.RecordCount = 1 Then
                    If InStr(ObjectOwner & ",", "," & grsObject!OWNER & "." & Split(strObject, ",")(i) & ",") = 0 Then
                        ObjectOwner = ObjectOwner & "," & grsObject!OWNER & "." & Split(strObject, ",")(i)
                    End If
                ElseIf grsObject.RecordCount > 1 Then
                    'ͬһ�����ж��������,��Ҫ��ѡ��
                    Set frmSelOwner.rsObject = grsObject
                    If frmParent Is Nothing Then
                        frmSelOwner.Show 1
                    Else
                        frmSelOwner.Show 1, frmParent
                    End If
                    If gblnOK Then
                        With frmSelOwner.lvw.SelectedItem
                            If InStr(ObjectOwner & ",", "," & .Text & "." & Split(strObject, ",")(i) & ",") = 0 Then
                                ObjectOwner = ObjectOwner & "," & .Text & "." & Split(strObject, ",")(i)
                            End If
                        End With
                        Unload frmSelOwner
                    Else
                        'ȡ��ѡ��,Ҳ����ȡ������(���ó���),���ؿ�
                        ObjectOwner = "ȡ��": Exit Function
                    End If
                End If
            End If
        End If
    Next
    If ObjectOwner <> "" Then ObjectOwner = Mid(ObjectOwner, 2)
End Function

Public Function SQLOwner(ByVal strSQL As String, strOwner As String) As String
'���ܣ���SQL����滻�ɴ����������ߵ���ʽ
'������strSQL=ԭʼSQL���,strOwner=���������ߴ�,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼"
'���أ����ʶ������������ǰ׺��SQL���
'˵����1.����������ֱ��ִ���û�SQL���,������Ҫ��Ȩ�����˽��ͬ��ʡ�
'      2.�Ա������ֶ�����ͬ���ֶ���û�д������,������
    Dim i As Integer, j As Integer
    Dim intLoc As Integer, blnDo As Boolean
    
    '�����ֻ�ÿո���
    strSQL = UCase(SpaceSQL(strSQL))
    
    For i = 0 To UBound(Split(strOwner, ","))
        '����ѭ��ȷ�Ϸ�ʽ,ȷ���滻���Ǳ���,������������䲿�ݻ򱻰��������������еĲ���
        j = 0 '��ǰ��ʼ����λ��
        Do
            j = j + 1
            intLoc = InStr(j, strSQL, Split(Split(strOwner, ",")(i), ".")(1))
            If intLoc > 12 Then '������"SELECT FROM "
                '�������������ǰ׺�Ĳ��滻
                blnDo = True
                '�ұ��Կո�","�š������Ž���
                blnDo = blnDo And (InStr(",) ", Mid(strSQL, intLoc + Len(Split(Split(strOwner, ",")(i), ".")(1)), 1)) > 0)
                '�����Ϊ","�Ż�"FROM "
                blnDo = blnDo And (Mid(strSQL, intLoc - 1, 1) = "," Or Mid(strSQL, intLoc - 5, 5) = "FROM ")
                If blnDo Then
                    strSQL = Left(strSQL, intLoc - 1) & _
                        Replace(strSQL, Split(Split(strOwner, ",")(i), ".")(1), Split(strOwner, ",")(i), intLoc, 1)
                    j = intLoc + Len(Split(strOwner, ",")(i))
                End If
            End If
        Loop Until j >= Len(strSQL)
    Next
    SQLOwner = strSQL
End Function

Public Function InDesign() As Boolean
'���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function GetDBUser() As String
'���ܣ���ȡ��ǰ��¼���ݿ��û���
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
        
    On Error GoTo errH
        
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then Exit Function
    If InStr(UCase(gcnOracle.ConnectionString), "USER ID=") > 0 Then
        For i = 0 To UBound(Split(UCase(gcnOracle.ConnectionString), ";"))
            If Split(UCase(gcnOracle.ConnectionString), ";")(i) Like "USER ID=*" Then
                GetDBUser = Trim(Split(Split(UCase(gcnOracle.ConnectionString), ";")(i), "=")(1))
                Exit For
            End If
        Next
    Else
        strSQL = "Select User From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ǰ��¼���ݿ��û���")
        If Not rsTmp.EOF Then GetDBUser = rsTmp!USER
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoSizeCol(lvw As Object)
'���ܣ������Զ�ListView��ǰ�����Զ��������п��
'������blnByHead=�Ƿ���ͷ�ı�����,Col=ָ���л���������(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ����洰�弰���и��ֿؼ���״̬
'������objForm:Ҫ����Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    
    Dim objThis As Object
    Dim strTmp As String
    Dim i As Integer, blnDo As Boolean
    
    On Error Resume Next
    If Not gfrmMain Is Nothing Then Call gfrmMain.Shut����(objForm)
    On Error GoTo 0
    
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "1") = "0" Then
        Call DelWinState(objForm, strProjectName, strUserDef)
        SaveWinState = True: Exit Function
    End If
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '���洰��״̬��λ�á���С
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form", "״̬", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form", "״̬", 0
            Case 2
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form", "״̬", objForm.WindowState
        End Select
    End With
    
    '������ֿؼ��ĸ���״̬
    For Each objThis In objForm.Controls
        strTmp = ""
        On Error Resume Next
        If UCase(TypeName(objThis)) = UCase("Menu") Then
            If objThis.Caption Like "��׼��ť*" Or _
                objThis.Caption Like "�ı���ǩ*" Or _
                objThis.Caption Like "״̬��*" Or _
                UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                '����˵��ĸ�ѡ
                strTmp = objThis.Checked & "," & objThis.Enabled
            Else
                strTmp = ""
            End If
        ElseIf (UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Or _
            UCase(TypeName(objThis)) = UCase("StatusBar") Or _
            UCase(TypeName(objThis)) = UCase("Toolbar") Or _
            UCase(TypeName(objThis)) = UCase("Coolbar")) And objForm.Visible Then

            blnDo = True
            If UCase(TypeName(objThis)) = UCase("Toolbar") Or UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Then
                If TypeName(objThis.Container) = "PictureBox" Then blnDo = False
            End If
            'Left,Top,Width��Height,Visible
            strTmp = strTmp & "," & objThis.Left
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Top
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Width
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Height
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            If blnDo Then
                strTmp = strTmp & "," & objThis.Visible
                If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            Else
                strTmp = strTmp & ",-32767"
            End If
            strTmp = Mid(strTmp, 2)
        End If
        If strTmp <> "" Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "״̬", strTmp
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("Toolbar")
                If objThis.Buttons.Count > 0 Then
                    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "�ı�", IIf(objThis.Buttons(1).Caption <> "", 1, objThis.ButtonHeight)
                End If
            Case UCase("ListView")
                SaveListViewState objThis, strProjectName & objForm.Name & strUserDef
            Case UCase("CoolBar")
                strTmp = ""
                For i = 1 To objThis.Bands.Count
                    strTmp = strTmp & "," & objThis.Bands(i).NewRow
                Next
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "����", Mid(strTmp, 2)
                
                strTmp = ""
                For i = 1 To objThis.Bands.Count
                    strTmp = strTmp & "," & objThis.Bands(i).Visible
                Next
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "�ɼ���", Mid(strTmp, 2)
        End Select
    Next
    SaveWinState = True
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ��ָ������״̬�����󶥱߽糬��ʱ�����Զ�����Ϊ0
'������objForm:Ҫ�ָ��Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
   
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object
    Dim blnDo As Boolean
    Dim strSave As String
    Dim strOEM As String
    
    On Error Resume Next
    
    If Not gfrmMain Is Nothing Then Call gfrmMain.Show����(objForm)
    
    blnDo = (GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0") = "1")
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '�ָ������״̬��λ�á���С
    If UCase(objForm.Name) = UCase("frmReport") _
        Or UCase(objForm.Name) = UCase("frmPreview") _
            Or UCase(objForm.Name) = UCase("frmDesign") Then
        strTmp = "2" '���ⴰ���ʼ���
    Else
        strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    End If
    If blnDo Then
        strSave = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form", "״̬", "")
        RestoreWinState = (strSave <> "")
        If strSave = "" Then strSave = strTmp
        aryInfo = Split(strSave, ",")
    Else
        aryInfo = Split(strTmp, ",")
    End If
    With objForm
        .WindowState = aryInfo(0)
        If UBound(aryInfo) = 4 Then
            .Left = IIf(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIf(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIf(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIf(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    '�ָ������и��ֿؼ��ĸ���״̬
    For Each objThis In objForm.Controls
        
        On Error Resume Next
        
        If blnDo Then
            strTmp = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '����˵��ĸ�ѡ
                If objThis.Caption Like "��׼��ť*" Or _
                    objThis.Caption Like "�ı���ǩ*" Or _
                    objThis.Caption Like "״̬��*" Or _
                    UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "״̬", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "״̬", "")
                If strTmp <> "" Then
                    'Left,Top,Width��Height,Visible
                    If UBound(Split(strTmp, ",")) = 4 Then
                        If Split(strTmp, ",")(0) <> "-32767" Then objThis.Left = Split(strTmp, ",")(0)
                        If Split(strTmp, ",")(1) <> "-32767" Then objThis.Top = Split(strTmp, ",")(1)
                        If Split(strTmp, ",")(2) <> "-32767" Then objThis.Width = Split(strTmp, ",")(2)
                        If Split(strTmp, ",")(3) <> "-32767" Then objThis.Height = Split(strTmp, ",")(3)
                        If Split(strTmp, ",")(4) <> "-32767" Then objThis.Visible = Split(strTmp, ",")(4)
                    End If
                End If
            End If
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("StatusBar")
                '״̬�����ñ�־
'                If zlRegInfo("��Ȩ����") <> "1" Then
'                    If objThis.Panels(1).Bevel = sbrRaised Then
'                        objThis.Panels(1).Text = ""
'                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
'                        objThis.Panels(1).ToolTipText = ""
'                        objThis.Height = 360
'                    End If
'                Else
                    If objThis.Panels(1).Bevel = sbrRaised Then
                        strTmp = zlRegInfo("��Ʒ����")
                        If strTmp <> "-" Then
                            objThis.Panels(1).Text = strTmp & "���"
                            '����״̬��ͼ���OEM����
                            If strTmp = "����" Then
                                If zlRegInfo("��Ȩ����") <> "1" Then
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
                                    objThis.Panels(1).Text = ""
                                Else
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                            Else
                                strOEM = GetOEM(strTmp)
                                Set objThis.Panels(1).Picture = LoadCustomPicture(strOEM)
                                If Err <> 0 Then
                                    Err.Clear
                                Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                                If zlRegInfo("��Ȩ����") <> "1" Then objThis.Panels(1).Text = strTmp & "(����)"
                            End If
                            objThis.Panels(1).ToolTipText = ""
                            objThis.Height = 360
                        End If
                    End If
'                End If
            Case UCase("Menu")
                If UCase(objThis.Name) = UCase("mnuHelpWeb") Then
                    'WEB�ϵ�����
                    strTmp = zlRegInfo("֧���̼���")
                    If strTmp <> "-" Then
                        objThis.Caption = "&WEB�ϵ�" & strTmp
                    End If
                ElseIf UCase(objThis.Name) = UCase("mnuHelpWebHome") Then
                    '������ҳ
                    strTmp = zlRegInfo("֧���̼���")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "��ҳ(&H)"
                    End If
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.Count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "�ı�", 1)
                        For i = 1 To objThis.Buttons.Count
                            objThis.Buttons(i).Caption = IIf(strTmp = 1, objThis.Buttons(i).Tag, "")
                        Next
                    End If
                End If
            Case UCase("ListView")
                If blnDo Then
                    RestoreListViewState objThis, strProjectName & objForm.Name & strUserDef
                End If
            Case UCase("CoolBar")
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "����", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "�ɼ���", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).Visible = Split(strTmp, ",")(i)
                        Next
                    End If
                End If
        End Select
    Next
End Function

Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim i As Integer, strTmp As String
        
    On Error Resume Next
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & "���", "")
    If UBound(Split(strTmp, ",")) >= 0 Then
        For i = 0 To objThis.Cols - 1
            If objThis.ColWidth(i) > 0 Then
                objThis.ColWidth(i) = Split(strTmp, ",")(i)
            End If
        Next
        RestoreFlexState = True
    End If
End Function

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strTmp As String, i As Integer
        
    On Error Resume Next
    
    strTmp = ""
    For i = 0 To objThis.Cols - 1
        strTmp = strTmp & "," & objThis.ColWidth(i)
    Next
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & "���", Mid(strTmp, 2)
End Sub

Public Sub SaveListViewState(objLvw As Object, ByVal strForm As String)
'���ܣ�����ListView�ĸ�������
'������objLvw=ListView����,strForm=����ؼ���
'˵������ͼ��ʽ���п���λ�á��б��⡢�ж��롢����
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String
    Dim strAlign As String
    
    For lngCol = 1 To objLvw.ColumnHeaders.Count
        strWidth = strWidth & "," & objLvw.ColumnHeaders(lngCol).Width
        strPosition = strPosition & "," & objLvw.ColumnHeaders(lngCol).Position
        strText = strText & "," & objLvw.ColumnHeaders(lngCol).Text
        strAlign = strAlign & "," & objLvw.ColumnHeaders(lngCol).Alignment
    Next
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "��ͼ", objLvw.View
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "���", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "λ��", Mid(strPosition, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "����", Mid(strText, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "����", Mid(strAlign, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "����", objLvw.SortKey & "," & objLvw.SortOrder & "," & objLvw.Sorted
End Sub

Public Sub RestoreListViewState(objLvw As Object, ByVal strForm As String)
'���ܣ��ָ�ListView�ĸ�������
'������objLvw=ListView����,strForm=����ؼ���
'˵������ͼ��ʽ���п���λ�á��б��⡢�ж��롢����
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String, varText As Variant
    Dim strAlign As String
    Dim strSort As String
    
    On Error Resume Next
    
    '��ͼȱʡ���ֳ�ʼֵ
    lngCol = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "��ͼ", -1)
    If lngCol <> -1 Then objLvw.View = lngCol
    
    strWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "���")
    strPosition = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "λ��")
    strAlign = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "����")
    For lngCol = 1 To objLvw.ColumnHeaders.Count
        '��ȱʡ�ؼ���Ϊ"_" & �б���
        objLvw.ColumnHeaders(lngCol).Key = "_" & objLvw.ColumnHeaders(lngCol).Text
        If strWidth <> "" Then objLvw.ColumnHeaders(lngCol).Width = Split(strWidth, ",")(lngCol - 1)
        If strPosition <> "" Then objLvw.ColumnHeaders(lngCol).Position = Split(strPosition, ",")(lngCol - 1)
        If strAlign <> "" Then objLvw.ColumnHeaders(lngCol).Alignment = Split(strAlign, ",")(lngCol - 1)
    Next
    
    '��������
    strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.Name & "����")
    If strSort <> "" Then
        objLvw.SortKey = Split(strSort, ",")(0)
        objLvw.SortOrder = Split(strSort, ",")(1)
        objLvw.Sorted = Split(strSort, ",")(2)
    End If
End Sub

Public Function DelWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ�ɾ��������Ի�����ֵ
'������objForm:Ҫ�ָ��Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    Dim strProject As String
    Dim lngR As Long
    Dim objThis As Object
    
    strProject = strProjectName
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    For Each objThis In objForm.Controls
        lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis) & Chr(0))
        If lngR <> 0 And lngR <> 2 Then Exit Function
    Next
    
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form" & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    
    DelWinState = True
End Function

Public Function LoadCustomPicture(strID As String, Optional strFormat As String = "GIF") As StdPicture
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function MatchIndex(ByVal cbo As Object, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������cbo.Hwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> cbo.hwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = cbo.hwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then
            cbo.Text = strFind
            cbo.SelStart = Len(cbo.Text)
        End If
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function RemoveOrderBy(ByVal str As String) As String
'���ܣ���SQL���������Order by ���ȥ��
    Dim i As Integer, intMax As Integer
    Dim strTmp As String
    
    strTmp = UCase(str): intMax = -1
    Do While strTmp Like UCase("*ORDER BY*")
        i = InStr(UCase(strTmp), "ORDER BY")
        If i > intMax Then intMax = i
        strTmp = Left(strTmp, i - 1) & "12345678" & Mid(strTmp, i + 8)
    Loop
    If intMax <> -1 Then
        RemoveOrderBy = Left(str, intMax - 1)
    Else
        RemoveOrderBy = str
    End If
End Function

Public Function GetDefaultValue(strSQL As String, strFld As String) As String
'���ܣ����ݲ���ѡ����SQL���壬������ʾ�ֶμ����ֶε�ֵ
'���أ���ʾֵ|��ֵ|��¼��
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim strShow As String, strBand As String
    Dim strSQLT As String
    
    On Error GoTo errH
    
    strSQLT = Replace(RemoveNote(strSQL), "[*]", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLT, "ѡ����SQL����")
    If Not rsTmp.EOF Then
        For i = 0 To UBound(Split(strFld, "|"))
            strTmp = Split(strFld, "|")(i)
            If Split(strTmp, ",")(2) Like "*&D*" Then
                strShow = IIf(IsNull(rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value), "", rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value)
            End If
            If Split(strTmp, ",")(2) Like "*&B*" Then
                strBand = IIf(IsNull(rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value), "", rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value)
            End If
        Next
    End If
    If strShow <> "" Or strBand <> "" Then
        GetDefaultValue = strShow & "|" & strBand & "|" & rsTmp.RecordCount
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, msg, wp, lp)
End Function

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function RemoveNote(ByVal strSQL As String) As String
'���ܣ��Ƴ�SQL����е�ע��
    Dim strTmp As String, i As Integer
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbCrLf, vbLf)
    strSQL = Replace(strSQL, vbCr, vbLf)
    strSQL = Replace(strSQL, vbLf & vbLf, vbLf)
    
    For i = 0 To UBound(Split(strSQL, vbLf))
        If Not Trim(Split(strSQL, vbLf)(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & Split(strSQL, vbLf)(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function

Public Function ShowHelpFunc(SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
'��ʾ��������
'SHwnd:���봰�ھ��(��Ϊ��������)
'htmName:��ӳ��CHM�е�htm�ļ�����
'Sys:ϵͳ,0:��������;1:zlhis
    Dim Path As String
    Dim strSave As String
    
    On Error GoTo ShowHelpErr
    
    ShowHelpFunc = False
    strSave = String(200, Chr$(0))
    If Sys = 0 Then
        Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
        If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
        Call Htmlhelp(SHwnd, Path, &H0, "zlreport\" & htmName & ".htm")
    Else
        If Mid(UCase(htmName), 5, 6) = "INSIDE" Then
            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
            Call Htmlhelp(SHwnd, Path, &H0, "zlreport\report.htm")
        Else
            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9app" & Trim(Format(Sys)) & ".chm"
            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
            strSave = "zl9app" & Trim(Format(Sys)) & "rpt\" & htmName & ".htm"
            Call Htmlhelp(SHwnd, Path, &H0, strSave)
        End If
    End If
    ShowHelpFunc = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Sub SetColWidth(msh As Control, objForm As Object)
'���ܣ��Զ���������п�,����С�ʺ�Ϊ׼
    Dim arrWidth() As Long
    Dim i As Integer, j As Integer
    
    ReDim arrWidth(msh.Cols - 1)
    
    msh.Redraw = False
    Set objForm.Font = msh.Font
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 Then
            For j = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1) To msh.Rows - 1
                If objForm.TextWidth(msh.TextMatrix(j, i) & "ab") + 45 > arrWidth(i) Then
                    arrWidth(i) = objForm.TextWidth(msh.TextMatrix(j, i) & "AB") + 45
                End If
            Next
        End If
    Next
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 Then msh.ColWidth(i) = arrWidth(i)
    Next
    msh.Redraw = True
End Sub

Public Function HaveDBA() As Boolean
'���ܣ��жϵ�ǰ�û��Ƿ����DBAȨ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From Session_Roles Where Role='DBA'"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�жϵ�ǰ�û�DBAȨ��")
    HaveDBA = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFunSource(strOwner As String, strFunc As String) As String
'���ܣ���ȡָ��������Դ����
'˵����1.���صĺ����ı��϶���"FUNCTION xxxxx"��ͷ
'      2.���������������ݿ�����vbLf����,vbcf��ת�����˿ո�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strText As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select * From All_Source Where TYPE='FUNCTION' And Upper(Owner)=Upper('" & strOwner & "') And Upper(Name)=Upper('" & strFunc & "') Order by Line"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ��������Դ����")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIf(IsNull(rsTmp!Text), "", rsTmp!Text)
            strText = strText & strTmp
            rsTmp.MoveNext
        Next
    End If
    GetFunSource = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReplaceName(ByVal strCode As String, strOld As String, strNew As String) As String
'���ܣ��ں��������н��������滻���µĺ�����
'˵����Ϊ�˲��ı亯������Ĵ�Сд,�����д˺���
'˵�������������������ݿ�����vbLf����,vbcf��ת�����˿ո�
    Dim i As Integer, strText As String
    Dim arrText() As String, strTmp As String
    
    arrText = Split(strCode, vbLf)
    For i = 0 To UBound(arrText)
        strTmp = arrText(i)
        If UCase(strTmp) Like UCase("*" & strOld & "*") Then
            strTmp = Replace(UCase(strTmp), UCase(strOld), UCase(strNew))
        End If
        strText = strText & vbCrLf & strTmp
    Next
    ReplaceName = Mid(strText, 3)
End Function

Public Function CheckParPrivs(ϵͳ As Long, ������ As String, ������ As Integer) As String
'���ܣ�����Ƿ����ָ�������е�ѡ��������Ĳ�ѯȨ��
'������������=���ڼ���ڸ����������Ƿ����Ȩ��
'���أ����������²�����Ȩ�޵Ķ���,��"���ű�,��Ա��"
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim arrObj() As String, i As Integer, j As Integer
    Dim strOwner As String, strObj As String
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFuncpars Where ���� is Not NULL And ϵͳ=" & ϵͳ & " And ������=" & ������
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯȨ��")
    
    For i = 1 To rsTmp.RecordCount
        strObj = Replace(rsTmp!����, "|", ",")
        If Left(strObj, 1) = "," Then strObj = Mid(strObj, 2)
        If Right(strObj, 1) = "," Then strObj = Mid(strObj, 1, Len(strObj) - 1)
        arrObj = Split(strObj, ",")
        For j = 0 To UBound(arrObj)
            strOwner = Split(arrObj(j), ".")(0)
            strObj = Split(arrObj(j), ".")(1)
            
            If gblnDBA Then
                grsObject.Filter = "Object_Type<>'FUNCTION' And Object_Name='" & UCase(strObj) & "'"
            Else
                grsObject.Filter = "OWNER='" & UCase(������) & "' And Object_Type<>'FUNCTION' And Object_Name='" & UCase(strObj) & "'"
            End If
            If grsObject.EOF Then
                CheckParPrivs = CheckParPrivs & "," & strObj
            End If
        Next
        rsTmp.MoveNext
    Next
    CheckParPrivs = Mid(CheckParPrivs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadFuncPars(ϵͳ As Long, ������ As Integer) As FuncPars
'���ܣ���ȡָ�������Ĳ�����
    Dim strSQL As String, objPars As New FuncPars
    Dim rsTmp As New ADODB.Recordset, i As Integer
    
    Set ReadFuncPars = New FuncPars
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFuncPars Where ϵͳ=" & ϵͳ & " And ������=" & ������ & " Order by ������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ�������Ĳ�����")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With rsTmp
                objPars.Add IIf(IsNull(!����), "", !����), !������, !������, _
                    IIf(IsNull(!������), "", !������), !����, IIf(IsNull(!ȱʡֵ), "", !ȱʡֵ), _
                    IIf(IsNull(!��ʽ), 0, !��ʽ), IIf(IsNull(!ֵ�б�), "", !ֵ�б�), _
                    IIf(IsNull(!����SQL), "", !����SQL), IIf(IsNull(!��ϸSQL), "", !��ϸSQL), _
                    IIf(IsNull(!�����ֶ�), "", !�����ֶ�), IIf(IsNull(!��ϸ�ֶ�), "", !��ϸ�ֶ�), _
                    IIf(IsNull(!����), "", !����), "_" & !������
            End With
            rsTmp.MoveNext
        Next
    End If
    Set ReadFuncPars = objPars
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFuncPars(ByVal strCode As String) As String
'���ܣ��Ӻ��������л�ȡ��������
'������strCode=��������,���ٴӲ�������"("��ʼ��
'���أ�"������,��������;...",��"NO_IN,VarChar2;ID_IN,Number;..."
'˵����1.��Ϊ������������,��������Ҳ������"���ű�.����%Type"����ʽ
'      2.����ֵ�в����ֲ�����IN/OUT���͡�
    Dim strTmp As String, i As Integer, j As Integer
    Dim blnStart As Boolean, arrPars() As String
    Dim strPar As String, arrOne() As String
    
    '�Ƴ�ע��
    strCode = RemoveNote(strCode)
    
    '���ÿո���
    strCode = Replace(strCode, vbTab, " ")
    strCode = Replace(strCode, vbCr, " ")
    strCode = Replace(strCode, vbLf, " ")
    
    '���Begin�ؼ��ֵĿ�ʼλ��:Begin���ᵥ����Ϊ������,������
    strTmp = "": blnStart = False: j = 0
    For i = 1 To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            blnStart = True
            strTmp = strTmp & Mid(strCode, i, 1)
        ElseIf blnStart Then
            blnStart = False
            If UCase(strTmp) = "BEGIN" Then
                j = i - Len("Begin")
                Exit For
            End If
            strTmp = ""
        End If
    Next
    If j = 0 Then Exit Function
    
    'Beginǰ��Ĵ���
    strCode = Trim(Left(strCode, j - 1))
    If InStr(strCode, "(") = 0 Then Exit Function
    '������()֮��Ĳ�������
    strCode = Trim(Mid(strCode, InStr(strCode, "(") + 1))
    strCode = Trim(Left(strCode, InStr(strCode, ")") - 1))
    
    '�����ֲ�����ʹ��"(x),(x,y)"
    If IsNumeric(Trim(strCode)) Then Exit Function
    arrPars = Split(strCode, ",") '������","�ż��
    For i = 0 To UBound(arrPars)
        If IsNumeric(Trim(arrPars(i))) Then Exit Function
    Next
    
    '�ֽ����:���Բ���IN.OUT
    For i = 0 To UBound(arrPars)
        arrOne = Split(Trim(arrPars(i)), " ")
        For j = 0 To UBound(arrOne)
            If j = 0 Then
                strPar = strPar & ";" & Trim(arrOne(j))
            ElseIf InStr(UCase(Trim(arrOne(j))), "CHAR") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "DATE") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "NUMBER") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "%TYPE") > 0 Then
                strPar = strPar & "," & Trim(arrOne(j))
                Exit For
            End If
        Next
    Next
    GetFuncPars = Mid(strPar, 2)
End Function

Public Function GetLenStr(str As String, lngW As Long, objBase As Object) As String
'���ܣ�����ָ���ĳ��Ƚ�ȡ�ַ���
    Dim lngTmp As Long, i As Integer
    
    For i = 1 To Len(str)
        lngTmp = lngTmp + objBase.TextWidth(Mid(str, i, 1))
        If lngTmp <= lngW Then
            GetLenStr = GetLenStr & Mid(str, i, 1)
        Else
            Exit For
        End If
    Next
    If GetLenStr <> str Then
        GetLenStr = Left(GetLenStr, Len(GetLenStr) - 1) & ".."
    End If
End Function

Public Function GetParVBMacro(str As String) As String
'����:�������������,������ת�����VB����ֵ
    Dim curDate As Date
    
    If InStr(str, "&") = 0 Then GetParVBMacro = str: Exit Function
    
    curDate = Currentdate
    Select Case str
        Case "&��ǰ����"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd")
        Case "&��ǰ����ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        Case "&ǰһ������"
            GetParVBMacro = Format(curDate - 7, "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -3, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(curDate + 7, "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 3, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd")
        Case "&���쿪ʼʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 00:00:00")
        Case "&�������ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&ǰһ��ͬʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd HH:mm:ss")
        Case "&��һ��ͬʱ��"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd HH:mm:ss")
        Case "&���³�ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59")
    End Select
End Function

Public Function ExeFunction(strOwner As String, strFunc As String, strPars As String, objPars As FuncPars) As String
'���ܣ�ִ��һ������
'������strOwner=������������
'      strFunc=������
'      strPars=�����Ĳ���������,��"NO_IN,Varchar;...",˳���뺯������һ��
'      objPars=�����Ĳ���ֵ����,��ǰֵ�����"ȱʡֵ"��
'���أ���ȷ=����ֵ������=������Ϣ,��"ERROR"��ͷ
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim arrPar() As String, i As Integer, j As Integer
    Dim StrName As String, strType As String
    On Error GoTo errH
    
    If strPars = "" Then
        strSQL = "Select " & strOwner & "." & strFunc & " as ����ֵ From Dual"
    Else
        arrPar = Split(strPars, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            strType = Split(arrPar(i), ",")(1)
            For j = 1 To objPars.Count
                If UCase(objPars(j).����) = UCase(StrName) Then Exit For
            Next
            If j <= objPars.Count Then
                'ȡ�����ֵ
                If UCase(strType) Like "*NUMBER*" Then
                    If objPars(j).ȱʡֵ = "" Then
                        strSQL = strSQL & ",NULL"
                    Else
                        strSQL = strSQL & "," & Val(objPars(j).ȱʡֵ)
                    End If
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'" & objPars(j).ȱʡֵ & "'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If IsDate(objPars(j).ȱʡֵ) Then
                        If Not (#1/1/3000# - CDate(objPars(j).ȱʡֵ)) Like "*.*" Then
                            strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                        Else
                            strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                    Else
                        strSQL = strSQL & ",NULL"
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            Else
                '�����͸�ȱʡֵ
                If UCase(strType) Like "*NUMBER*" Then
                    strSQL = strSQL & ",1"
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'A'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    strSQL = strSQL & ",To_DATE('" & Format(Currentdate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                Else
                    strSQL = strSQL & ",NULL"
                End If
            End If
        Next
        strSQL = "Select " & strOwner & "." & strFunc & "(" & Mid(strSQL, 2) & ") as ����ֵ From Dual"
    End If
    
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ִ��һ������")
    
    If Err.Number = 0 Then
        ExeFunction = IIf(IsNull(rsTmp!����ֵ), "", rsTmp!����ֵ)
    Else
        ExeFunction = "ERROR" & Err.Description
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetFunctionExp(strOwner As String, strFunc As String, strPars As String, objPars As FuncPars) As String
'���ܣ����غ�����ִ�й�ʽ
'������strOwner=������������
'      strFunc=������
'      strPars=�����Ĳ���������,��"NO_IN,Varchar;...",˳���뺯������һ��
'      objPars=�����Ĳ���ֵ����,��ǰֵ�����"ȱʡֵ"��
'˵�������ڶ�̬ʱ�����,���ز�����ʽΪ"[zlBeginTime]","[zlEndTime]"
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim arrPar() As String, i As Integer, j As Integer
    Dim StrName As String, strType As String
    
    If strPars = "" Then
        strSQL = strOwner & "." & strFunc
    Else
        arrPar = Split(strPars, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            strType = Split(arrPar(i), ",")(1)
            
            For j = 1 To objPars.Count
                If UCase(objPars(j).����) = UCase(StrName) Then Exit For
            Next
            
            If j <= objPars.Count Then
                'ȡ�����ֵ
                If UCase(strType) Like "*NUMBER*" Then
                    If objPars(j).ȱʡֵ = "" Then
                        strSQL = strSQL & ",NULL"
                    Else
                        strSQL = strSQL & "," & Val(objPars(j).ȱʡֵ)
                    End If
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'" & objPars(j).ȱʡֵ & "'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If UCase(StrName) = "ZLBEGINTIME" Or UCase(StrName) = "ZLENDTIME" Then
                        If IsDate(objPars(j).ȱʡֵ) Then
                            If Not (#1/1/3000# - CDate(objPars(j).ȱʡֵ)) Like "*.*" Then
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            Else
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                        Else
                            strSQL = strSQL & ",[" & StrName & "]"
                        End If
                    Else
                        If IsDate(objPars(j).ȱʡֵ) Then
                            If Not (#1/1/3000# - CDate(objPars(j).ȱʡֵ)) Like "*.*" Then
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            Else
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                        Else
                            strSQL = strSQL & ",NULL"
                        End If
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            Else
                '�����͸�ȱʡֵ
                If UCase(strType) Like "*NUMBER*" Then
                    strSQL = strSQL & ",1"
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'A'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If UCase(StrName) = "ZLBEGINTIME" Or UCase(StrName) = "ZLENDTIME" Then
                        strSQL = strSQL & ",[" & StrName & "]"
                    Else
                        strSQL = strSQL & ",To_DATE('" & Format(Currentdate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            End If
        Next
        strSQL = strOwner & "." & strFunc & "(" & Mid(strSQL, 2) & ")"
    End If
    GetFunctionExp = strSQL
End Function

Public Function GetFuncName(ByVal strCode As String) As String
'���ܣ����ݺ��������ȡ������
    Dim strTmp As String, blnStart As Boolean
    Dim i As Integer, j As Integer
    
    '�Ƴ�ע��
    strCode = RemoveNote(strCode)
    
    '���ÿո���
    strCode = Replace(strCode, vbTab, " ")
    strCode = Replace(strCode, vbCr, " ")
    strCode = Replace(strCode, vbLf, " ")
    
    '���Begin�ؼ��ֵĿ�ʼλ��:Begin���ᵥ����Ϊ������,������
    strTmp = "": blnStart = False: j = 0
    For i = 1 To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            blnStart = True
            strTmp = strTmp & Mid(strCode, i, 1)
        ElseIf blnStart Then
            blnStart = False
            If UCase(strTmp) = "BEGIN" Then
                j = i - Len("Begin")
                Exit For
            End If
            strTmp = ""
        End If
    Next
    
    'û��Begin,���������
    If j = 0 Then Exit Function
    
    'Beginǰ��Ĵ���
    strCode = Trim(Left(strCode, j - 1))
    
    'û��Function,���������
    j = InStr(UCase(strCode), "FUNCTION")
    If j = 0 Then Exit Function
    j = j + Len("FUNCTION")
    
    '��������ʼ
    For i = j To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            j = i: Exit For
        End If
    Next
    If i > Len(strCode) Then Exit Function
    
    'ȡ������
    strTmp = ""
    For i = j To Len(strCode)
        '���Բ�����"("����������
        If Mid(strCode, i, 1) = " " Or Mid(strCode, i, 1) = "(" Then Exit For
        strTmp = strTmp & Mid(strCode, i, 1)
    Next
    GetFuncName = strTmp
End Function

Public Function FuncOwnerName(ByVal strCode As String, StrName As String, strOwner As String) As String
'���ܣ��ں��������к�����ǰ����������
'������strName=������
    Dim i As Integer, strTmp As String
    
    i = InStr(UCase(strCode), UCase(StrName))
    If i = 0 Then
        strTmp = strCode
    Else
        strTmp = Left(strCode, i - 1) & strOwner & "." & Mid(strCode, i)
    End If
    FuncOwnerName = strTmp
End Function

Public Function GetBalndValue(strSQL As String, strFld As String, strVal As String) As String
'���ܣ����ݲ���ѡ����SQL���壬����ָ����ֵ����ʾ�ֶμ����ֶε�ֵ
'���أ���ʾֵ|��ֵ
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim strShowFld As String, strBandFld As String
    Dim strShow As String, strBand As String
    Dim strSQLT As String
    
    On Error GoTo errH
    
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then
            strShowFld = CStr(Split(strTmp, ",")(0))
        End If
        If Split(strTmp, ",")(2) Like "*&B*" Then
            strBandFld = CStr(Split(strTmp, ",")(0))
        End If
    Next
    strSQLT = Replace(RemoveNote(strSQL), "[*]", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLT, "���ݲ���ѡ����SQL����")
    
    Do While Not rsTmp.EOF
        strShow = IIf(IsNull(rsTmp.Fields(strShowFld).Value), "", rsTmp.Fields(strShowFld).Value)
        strBand = IIf(IsNull(rsTmp.Fields(strBandFld).Value), "", rsTmp.Fields(strBandFld).Value)
        If strBand = strVal Then Exit Do
        rsTmp.MoveNext
    Loop
    If strShow <> "" Or strBand <> "" Then GetBalndValue = strShow & "|" & strBand
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetDate(strVal As String) As Date
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & strVal & " as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
    GetDate = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SplitFunc(ByVal strExp As String, strOwner As String, strFunc As String, strPars As String)
'���ܣ����ݺ������ʽ�ֽ����
'������strExp="ZLCOST.ZL_FUN_GATHER([ZLBEGINTIME],[ZLENDTIME],102,2558)"
'���أ�strOwner=����������,strFunc=������,strPars=��"|"�����ԭ����,��"[ZLBEGINTIME]|[ZLENDTIME]|102|'����'"
    Dim i As Integer, intA As Integer, intB As Integer, intSign As Single

    If Trim(strExp) = "" Then Exit Sub
    If Not UCase(strExp) Like "*.ZL_FUN_*" Then Exit Sub
    
    strOwner = UCase(Left(strExp, InStr(strExp, ".") - 1))
    strExp = Mid(strExp, InStr(strExp, ".") + 1)
    If InStr(strExp, "(") = 0 Then
        strFunc = strExp
        strPars = ""
    Else
        strFunc = Left(strExp, InStr(strExp, "(") - 1)
        strExp = Mid(strExp, InStr(strExp, "(") + 1)
        strExp = Left(strExp, Len(strExp) - 1)
        
        intA = 0: intB = 0: intSign = 1
        For i = 1 To Len(strExp)
            
            If Mid(strExp, i, 1) = "(" Then
                intA = intA + 1
            ElseIf Mid(strExp, i, 1) = ")" Then
                intA = intA - 1
            ElseIf Mid(strExp, i, 1) = "'" Then
                intB = intB + intSign
                intSign = -1 * intSign
            End If
            
            If Mid(strExp, i, 1) = "," And intA = 0 And intB = 0 Then
                strPars = strPars & "|"
            Else
                strPars = strPars & Mid(strExp, i, 1)
            End If
        Next
    End If
End Sub

Public Function GetFuncSys(strOwner As String, strFunc As String) As Long
'���ܣ���ȡ���ݿ⺯������ϵͳ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ϵͳ From zlFunctions Where Upper(������)='" & UCase(strFunc) & "' And ϵͳ IN(Select ��� From zlSystems Where Upper(������)='" & UCase(strOwner) & "')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݿ⺯������ϵͳ")
    If Not rsTmp.EOF Then GetFuncSys = rsTmp!ϵͳ
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SpaceSQL(ByVal strSQL As String) As String
'���ܣ���SQL���任ΪֻΪ�ո�������ʽ,�Ա��ڷ���
    Dim i As Long, j As Long, lngB As Long, lngE As Long
    Dim arrSeg() As Variant
                
    strSQL = Replace(strSQL, vbCr, " ")
    strSQL = Replace(strSQL, vbLf, " ")
    strSQL = Replace(strSQL, vbTab, " ")
    
    lngB = -1
    arrSeg = Array()
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "'" Then
            If lngB = -1 Then
                lngB = i
            Else
                ReDim Preserve arrSeg(UBound(arrSeg) + 1)
                arrSeg(UBound(arrSeg)) = lngB & "," & i
                lngB = -1
            End If
        End If
    Next
    If lngB = -1 Then
        For i = 0 To UBound(arrSeg)
            lngB = CLng(Split(arrSeg(i), ",")(0)) + 1
            lngE = CLng(Split(arrSeg(i), ",")(1)) - 1
            For j = lngB To lngE
                If Mid(strSQL, j, 1) = " " Then
                    strSQL = Left(strSQL, j - 1) & Chr(250) & Mid(strSQL, j + 1)
                End If
            Next
        Next
    End If
    
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    
    strSQL = Replace(strSQL, Chr(250), " ")
    
    strSQL = Replace(strSQL, " ,", ",")
    strSQL = Replace(strSQL, ", ", ",")
    SpaceSQL = strSQL
End Function

Public Function zlHomePage(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����룬������ҳ
    Dim strCode As String
    
    strCode = zlRegInfo("֧����URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����룬������̳
    Dim strCode As String
    
    'strCode = zlRegInfo("֧����BBS")
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Function zlMailTo(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����뷢�͵����ʼ�
    Dim strCode As String
    strCode = zlRegInfo("֧����MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function
