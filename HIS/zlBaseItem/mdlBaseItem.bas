Attribute VB_Name = "mdlBaseItem"
Option Explicit
Public gblnʹ����ҽ As Boolean
Public gbln������ҽ As Boolean
Public gstrҽ�۽ӿڱ�� As String
Public gbln����ҽ���շ���Ŀ As Boolean
Public gbln��������ۿ�  As Boolean
'��ҹ���
Public gobjPlugIn As Object
Public gblnMyStyle As Boolean
Public gstrMatchMode As String
Public gbytCode As Byte
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--����ϵͳ����
'����:27990
Private Type Ty_System_Para
     bytҩƷ������ʾ As Byte   'ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ����0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
     byt����ҩƷ��ʾ As Byte  '����ҩƷ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ����0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
End Type
Public gTy_System_Para As Ty_System_Para
Public gblnFeeKindCode As Boolean
Public gstrҩƷ�۸�ȼ� As String
Public gstr���ļ۸�ȼ� As String
Public gstr��ͨ�۸�ȼ� As String
'Windows���----------------------------------
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
  
Public Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public gstrLike As String
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'�趨һ�����岶����꣬���������������Ϣ�������ô���
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'ȡ����겶��
Public Declare Function ReleaseCapture Lib "user32" () As Long

'IP��ַ��ʽ���
Public Declare Function inet_addr Lib "ws2_32" (ByVal lpszAddress As String) As Long
Public Const INADDR_NONE = &HFFFFFFFF

'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const GCST_INVALIDCHAR = " '"    '�����������Ч�ַ�
Public gobjCustAcc As Object

Public Enum EditMode 'medit��ʽ  ȡֵΪ��0��������1���޸ģ�2�����ۣ�3��ִ�п��ҡ�4��������Ŀ��5�������޸�ִ�п���
    EditNew = 0
    EditModify = 1
    EditRaise = 2
    EditDept = 3
    EditSlave = 4
    EditCopy = 5
End Enum
Public gobjNurseIntegrate As Object  '���廤��ӿڶ���
Public gobjRIS As Object                    '����RIS�ӿڶ���
Public Enum RISBaseItemOper                 '����RIS�������ݲ������ͣ�1-������2-�޸ģ�3-ɾ��
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '����RIS�����������ͣ�3���û�(��Ա��
    Personnel = 3
End Enum

'������־ģ��
Private mobjFso As New FileSystemObject '�ļ�����

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub SetFormVisible(ByVal new_Hwnd As Long)
    '���ش��������С��ť
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 Or WS_SYSMENU Or &H20000
End Sub

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'���ܣ���ʼ�������ӿڲ���
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����ڵ���ʾ������
    Dim vRect As RECT, vPos As POINTAPI
    
    vPos = zlControl.GetCursorPosition
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Function MoveSpecialChar(ByVal strInputString As String) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intASC As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intASC = Asc(Mid(strText, n, 1))
        Select Case intASC
            Case 8, 9, 10, 13, 32
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

Public Sub �ı����(nodParent As Node, int��ȥ���� As Integer, str�������� As String)
'����:�ı������б���ڵ�ı����б����ֵ
'����:nodParent         Ҫ�ı�������ʼ�ڵ�
'     int��ȥ����       ��������ȥ����
'     str��������       ��������������

    Dim nod As Node
    '�����¼�ҲҪ�ı����
    If nodParent.Children > 0 Then
        Set nod = nodParent.Child
        Do While Not (nod Is Nothing)
            nod.Text = "��" & str�������� & Mid(nod.Text, int��ȥ���� + 2)
            �ı���� nod, int��ȥ����, str��������
            Set nod = nod.Next
        Loop
    End If
End Sub

Public Function GetRoot(ByVal nod As Node) As Node
'���ܣ���������ڵ�ĸ��ڵ�
    Dim nodTemp As Node
    
    If nod Is Nothing Then Exit Function
    Set nodTemp = nod
    Do Until nodTemp.Parent Is Nothing
        Set nodTemp = nodTemp.Parent
    Loop
    Set GetRoot = nodTemp
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = "-") As String
'������cmbTemp  ׼����ȡ���ݵ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        'ֱ�ӷ��������ַ���
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            'Բ��֮ǰ
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = "-")
'������cmbTemp  ׼�����õ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = cmbTemp.Text Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '�Ѿ��ҵ�
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Sub ���鱨��(frmParent As Form)
    MsgBox "�����в���ϵͳ����Ա����", vbInformation, gstrSysName
End Sub


Public Function GetPictureInfo(picTemp As StdPicture, Optional strBitmap As String = "") As String
'���һ��ͼƬ����Ϣ
    Dim hFile As Integer
    Dim FileHeader As BITMAPFILEHEADER
    Dim InfoHeader As BITMAPINFOHEADER
    
    If picTemp.Handle = 0 Then
        GetPictureInfo = "����Ƭ"
        Exit Function
    End If
    
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    
    If strBitmap = "" Then
        '������ʱ�ļ�
        strPath = Space(256): strFile = Space(256)
        GetTempPath 256, strPath
        strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
        
        GetTempFileName strPath, "pic", 0, strFile
        strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    
        SavePicture picTemp, strFile
    Else
        'ֱ��ʹ�������ļ�
        strFile = strBitmap
    End If
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
      Get #hFile, , FileHeader
      Get #hFile, , InfoHeader
    Close #hFile
    
    If strBitmap = "" Then
        'ɾ����ʱ�ļ�
        Kill strFile
    End If
    
    If InfoHeader.biBitCount > 8 Then
         GetPictureInfo = InfoHeader.biWidth & "��" & InfoHeader.biHeight & " " & InfoHeader.biBitCount & "λɫ"
    Else
         GetPictureInfo = InfoHeader.biWidth & "��" & InfoHeader.biHeight & " " & 2 ^ InfoHeader.biBitCount & "ɫ"
    End If
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Sub InitSystemPara()
    '����ȫ�ֲ���
    '-------------------------------------------------------------------------------------------------
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    '���������ʱ,���������Ŀʱ,��λ����������
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1"
    gstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gbln��������ۿ� = zlDatabase.GetPara(93, glngSys) = "1"
    '����:27990
    With gTy_System_Para
        .byt����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ")) '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
        .bytҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ"))  '��0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    End With
End Sub
Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����, ����, ���� From �շ���Ŀ���"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ����")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ϣ��
    '���:strMsgInfor-��ʾ��Ϣ
    '        blnYesNo-�Ƿ��ṩYES��NO��ť
    '����:
    '����:blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '����:���˺�
    '����:2010-08-27 16:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function MedicalTeamPatients(ByVal lngTeamID As Long, ByVal lngMemberID As Long) As String
'----------------------------------------------------------------------
'���ܣ� �г�ҽ��С��ҽ���Ĳ�����Ϣ
'������ lngTeamID: ҽ��С��ID
'       lngMemberID: ҽ��ID
'���أ� ������Ϣ�ַ���
'----------------------------------------------------------------------
    Dim strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.����id, a.סԺ��, a.��Ժ����, b.����" & vbNewLine & _
              "From ������ҳ a, ������Ϣ b " & vbNewLine & _
              "Where a.סԺҽʦ = (Select ����" & vbNewLine & _
              "              From ��Ա��" & vbNewLine & _
              "              Where ID = [2] And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)) And" & vbNewLine & _
              "      a.ҽ��С��id = [1] and a.����id=b.����id and a.��ҳid=b.��ҳid and b.��Ժ=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��С��ҽ��������Ϣ", lngTeamID, lngMemberID)
    With rsTmp
        For i = 1 To .RecordCount
            strMess = strMess & "������" & !���� & "��" & vbTab & _
                      "סԺ�ţ�" & IIF(IsNull(!סԺ��), "", !סԺ��) & "��" & vbTab & _
                      "���ţ�" & IIF(IsNull(!��Ժ����), "", !��Ժ����) & vbTab & vbNewLine
            .MoveNext
        Next
    End With
    MedicalTeamPatients = strMess
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDeptPermission(ByVal lngOperationID As Long, Optional ByVal lngDeptID As Long) As Boolean
'����: ��鲿��Ȩ��
'lngOperationID: Ҫ��������ԱID
'lngDeptID: Ҫ������Ա�Ĳ���ID
'����: True��Ȩ��, False��Ȩ��
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    If lngDeptID = 0 Then
        gstrSQL = "Select Count(*) Rec From ������Ա " & _
                  "Where ��Աid = [2] And [3] In (Select ����id From ������Ա Where ��Աid = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ա�Ĳ���Ȩ��", glngUserId, lngOperationID, lngDeptID)
    Else
        gstrSQL = "Select ID " & _
                  "From ���ű� " & _
                  "  Start With ID In (Select ����id From ������Ա Where ��Աid = [1]) " & _
                  "  Connect By Prior ID = �ϼ�id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ա�Ĳ���Ȩ��", glngUserId)
        Do While Not rsTmp.EOF
            If rsTmp!ID = lngDeptID Then
                CheckDeptPermission = True
                Exit Function
            End If
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End If
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetBillFormatRec(ByVal strReportCode As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������Ĵ�ӡ��ʽ
    '���:strReportCode-��������
    '����:�����ӡ��ʽ�ļ�¼��
    '����:���˺�
    '����:2015-06-10 11:43:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  " & _
    "   From Dual " & _
    "   Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1] " & _
    "   Order by ���"
    Set zlGetBillFormatRec = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strReportCode)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPriceGrade(ByRef strҩƷ�۸�ȼ� As String, _
    ByRef str���ļ۸�ȼ� As String, ByRef str��ͨ�۸�ȼ� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰվ��۸�ȼ�
    '���:
    '����:�۸�ȼ���ȡ�ɹ�����True�����򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    strҩƷ�۸�ȼ� = "": str���ļ۸�ȼ� = "": str��ͨ�۸�ȼ� = ""
    strSQL = "" & _
        "Select Max(Decode(b.�Ƿ�����ҩƷ, 1, �۸�ȼ�, Null)) As ҩƷ�ȼ�," & vbNewLine & _
        "       Max(Decode(b.�Ƿ���������, 1, �۸�ȼ�, Null)) As ���ĵȼ�," & vbNewLine & _
        "       Max(Decode(b.�Ƿ�������ͨ��Ŀ, 1, �۸�ȼ�, Null)) As ��ͨ�ȼ�" & vbNewLine & _
        "From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B" & vbNewLine & _
        "Where a.�۸�ȼ� = b.���� And a.���� = 0 And a.վ�� = [1]" & vbNewLine & _
        "      And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�۸�ȼ�", gstrNodeNo)
    If Not rsTemp.EOF Then
        strҩƷ�۸�ȼ� = Nvl(rsTemp!ҩƷ�ȼ�)
        str���ļ۸�ȼ� = Nvl(rsTemp!���ĵȼ�)
        str��ͨ�۸�ȼ� = Nvl(rsTemp!��ͨ�ȼ�)
    End If
    GetPriceGrade = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File

    If Val(OS.IniRead("LOG", "OPENLOG", App.Path & "\CONFIG.INI")) = 0 Then Exit Sub
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\��־*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    '3���ռ��Ƿ��㹻
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '�ռ䲻�㣬��д��־,����һ�������ļ�
        If Not mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\�ռ䲻��.txt", True)
        Exit Sub
    Else
        '��������ļ�
        If mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.DeleteFile(strLogPath & "\�ռ䲻��.txt", True)
    End If
    '4��д����־��
    strLogFile = strLogPath & "\��־" & Format(Now, "yyyyMMdd") & ".log"

    Call SaveLog(strLogFile, strLogTxt)

End Sub

Public Sub SaveLog(ByVal strFileName As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFileName) Then Call mobjFso.CreateTextFile(strFileName)
        Set objStream = mobjFso.OpenTextFile(strFileName, ForAppending)
        If strDate = "" Then
            strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
            objStream.WriteLine (strDate & Chr(&H9) & strInput)
        Else
            objStream.WriteLine (strInput)
        End If
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '��ȡʣ��ռ�
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Function FuncGetStr(ByVal strVal As String) As String
    strVal = Replace(strVal, vbTab, "")
    strVal = Replace(strVal, vbCrLf, "")
    strVal = Replace(strVal, Chr(10), "")
    strVal = Replace(strVal, "'", "''")
    strVal = Replace(strVal, " ", "")
    FuncGetStr = Trim(strVal)
End Function

Public Function IsPriceGradeEnabled() As Boolean
    '�Ƿ������˼۸�ȼ�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = "Select 1 From �շѼ۸�ȼ�Ӧ�� Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ������˼۸�ȼ�")
    IsPriceGradeEnabled = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsIPAddress(ByVal strAddress As String) As Boolean
'���ܣ��ж������Ip��ַ��ʽ�Ƿ�Ϸ�
    If inet_addr(strAddress) <> INADDR_NONE Then
        IsIPAddress = True
    Else
        IsIPAddress = False
    End If
End Function

Public Function InitNurseIntegrate(Optional blnMsg As Boolean = False) As Boolean
'�ж�������廤����Ϊ�վͳ�ʼ��
    If gobjNurseIntegrate Is Nothing Then
        On Error Resume Next
        Set gobjNurseIntegrate = CreateObject("zlNurseIntegrate.clsNurseIntegrate")
        If Not gobjNurseIntegrate Is Nothing Then
            If gobjNurseIntegrate.zlInitCommon(gcnOracle, gstrDbUser) = False Then
                Set gobjNurseIntegrate = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    If blnMsg = True And gobjNurseIntegrate Is Nothing Then
        MsgBox "���廤��ӿڲ�����zlNurseIntegrate  ����ʧ�ܣ�", vbInformation, gstrSysName
    End If
    InitNurseIntegrate = Not gobjNurseIntegrate Is Nothing
End Function
