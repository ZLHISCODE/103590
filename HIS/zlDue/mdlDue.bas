Attribute VB_Name = "mdlDue"
Option Explicit

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gstrAviPath As String
Public gstrVersion As String
Public gstrMatchMethod As String

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public gstrIme As String

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum mAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

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
Private Type POINTAPI
        X As Long
        Y As Long
End Type
 
'�л���ָ�������뷨��
'Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
'Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
'Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
'Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'��ȡָ�����뷨����Layout,����Ϊ0ʱ��ʾ��ǰ���뷨��
'Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'��ȡ��ǰ���뷨����Layout��
'Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'�������뷨Layout���������뷨�л������뷨�л�˳�����ǰͷ(������������Ч),flags����=KLF_REORDER
'Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type SystemParameter
    int���뷽ʽ As Integer
    bln���Ի���� As Boolean               'ʹ�ø��Ի����
    Para_���뷽ʽ As String             ''��1λ1-ȫ����ֻ�����,��2λ1-ȫ��ĸֻ�����,��HIS��������������
    bln����վ�� As Boolean      '�Ƿ����վ�����
End Type
Public Enum gС������
    g_���� = 0
    g_�ɱ���
    g_�ۼ�
    g_���
End Enum

Private Type m_С��λ
    ����С�� As Integer
    �ɱ���С�� As Integer
    ���ۼ�С�� As Integer
    ���С�� As Integer
End Type

Public g_С��λ�� As m_С��λ

'С����ʽ����
Public Type g_FmtString
    FM_���� As String
    FM_�ɱ��� As String
    FM_���ۼ� As String
    FM_��� As String
End Type

Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
Public gSystemPara As SystemParameter


'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48

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

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

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


Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo ErrHand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
ErrHand:
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = "" & rsTmp!����
        UserInfo.���� = "" & rsTmp!����
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'���º���û��ʹ�� by lesfeng 2009-12-2 �����Ż�
'Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
'    '������������ȡָ����ı����������󳤶�
'    '�������������ID������
'    '����������ɹ����� �¼�������; ���߷��� 0
'    Dim strSQL As String
'    Dim rsTemp As New ADODB.Recordset
'
'    Err = 0
'    On Error GoTo Error_Handle
'    If strID = "" Then
'        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID is null " & strWhere & " connect by prior id=�ϼ�id"
'        zldatabase.OpenRecordset rsTemp, strSQL, "��ȡָ����ı����������󳤶�"
'    Else
'        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with ID=[1] " & strWhere & " connect by prior id=�ϼ�id"
'        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡָ����ı����������󳤶�", CLng(strID))
'    End If
'
'    If rsTemp.EOF Then
'        GetDownCodeLength = 0
'    Else
'        GetDownCodeLength = rsTemp.Fields("LenCode").Value
'    End If
'    Exit Function
'Error_Handle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'    GetDownCodeLength = 0
'End Function

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID is null" & strWhere
        zlDatabase.OpenRecordset rsTemp, strSQL, "��ȡָ����ı����������󳤶�"
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID=[1]" & strWhere
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ����ı����������󳤶�", CLng(str�ϼ�ID))
    End If
    
    
    If rsTemp.EOF Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select ���� from " & strTableName & " where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ϼ�����", CLng(str�ϼ�ID))
    If rsTemp.EOF Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

'Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
'    '��������������ָ������ϼ�ID ��ȡ������������
'    '����������ϼ�ID,����
'    '����������ɹ����� ������; ���߷��� ��
'    Dim strSQL As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim intCode As Integer, StrCode As String, strAllCode As String
'    Dim intLength   As Integer
'    Err = 0
'    On Error GoTo Error_Handle
'    If str�ϼ�ID = "" Then
'        strSQL = "select max(to_number(����))+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
'        zldatabase.OpenRecordset rsTemp, strSQL, "����ָ������ϼ�ID ��ȡ������������"
'    Else
'        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=[1]" & strWhere
'        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "����ָ������ϼ�ID ��ȡ������������", CLng(str�ϼ�ID))
'    End If
'    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
'
'    If rsTemp.EOF Then
'        GetMaxLocalCode = ""
'        Exit Function
'    End If
'    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
'    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
'    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
'    Exit Function
'Error_Handle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'    GetMaxLocalCode = ""
'End Function

Public Function NextNo(intBillId As Integer) As Variant
    '------------------------------------------------------------------------------------
    '���ܣ������ض���������µ���ⵥ����,�������£�
    '       ���λȷ��ԭ��:
    '       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���
    '���أ�
    '------------------------------------------------------------------------------------
    Dim rsCtrl As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim vntNo As Variant        '��ȡ�ĺ�����м����
    Dim intYear, strYear As String      '��ȱ�־λ

RESTART:
    Err = 0
    On Error GoTo ErrHand
    
    With rsCtrl
        If .State = adStateOpen Then .Close
        .Open "Select C.��Ŀ���,C.��Ŀ����,C.������,C.�Զ���ȱ,C.��Ź���,sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillId, gcnOracle, adOpenKeyset, adLockOptimistic
        If .EOF Or .BOF Then
            NextNo = Null
            Exit Function
        End If
        intYear = Format(!Today, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        vntNo = IIf(IsNull(!������), "", !������)
        If Left(vntNo, 1) < strYear Then
            vntNo = strYear & "0000000"
        End If
        vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
        
        On Error Resume Next
        .Update "������", vntNo
        If Err <> 0 Then
            .CancelUpdate
            GoTo RESTART
        End If
        NextNo = vntNo
    End With
    Exit Function

ErrHand:
    Call ErrCenter
    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function

'Public Function BinTOHex(sString As String) As String
'    Dim lngLoop As Integer, lngTemp As Long, lngJLoop As Integer, lngTmp As Long
'    lngTemp = 0
'    For lngLoop = 1 To Len(sString)
'        If Mid(sString, lngLoop, 1) = "1" Then
'            lngTmp = 1
'            For lngJLoop = 0 To lngLoop - 2
'                lngTmp = lngTmp * 2
'            Next
'        Else
'            lngTmp = 0
'        End If
'        lngTemp = lngTemp + lngTmp
'    Next
'    BinTOHex = CStr(lngTemp)
'End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function CheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:20070101��2007-01-01
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgbox strTittle & "����Ϊ������,���飡"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgbox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡"
        Exit Function
    End If
    CheckIsDate = strKey
End Function

Public Sub zlChangeCode(ByVal strTableName As String, _
    ByVal lng�ϼ�id As Long, _
    ByVal txtUpCode As TextBox, _
    ByVal txtCode As TextBox, _
    Optional ByVal chkChangeCode As CheckBox = Nothing, _
    Optional ByVal strCaption As String = "")
    '------------------------------------------------------------------------------------
    '���ܣ�����ѡ����ϼ�ȷ����ǰ�ı��룬�����ϼ�����������ʾ����
    '������strTableName-���ڷ���ı���
    '      lng�ϼ�ID-ѡ����ϼ�
    '      TxtUpCode-��ʾ���ϼ��ı���
    '      TxtUpCode-��ʾ�ı����ı���
    '      chkChangeCode-�����Ƿ�ı�ԭ�����ݿ��е���ʷ����ѡ��ؼ�
    '      strCaption-���ô����Capiton
    'ע�⣺���б�����ID,�ϼ�id,����
    '------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intMaxCodeLen As Integer  'ȷ�������ʵ�ʳ���
    Err = 0: On Error GoTo ErrHand
    
   chkChangeCode.Value = 0
   chkChangeCode.Enabled = True
   
    If lng�ϼ�id = 0 Then
        txtUpCode.Text = ""
        gstrSQL = "select max(����) as ���� From " & strTableName & " Where �ϼ�ID is null "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
            
        With rsTemp
            intMaxCodeLen = .Fields("����").DefinedSize
            If IsNull(!����) Then
                txtCode.Text = "01"
                txtCode.MaxLength = intMaxCodeLen
                txtCode.Tag = txtCode.MaxLength
                chkChangeCode.Value = 1
                chkChangeCode.Enabled = False
            Else
                txtCode.MaxLength = Len(Trim(!����))
                txtCode.Tag = txtCode.MaxLength
                If !���� = String(txtCode.MaxLength, "9") Then
                    If txtCode.MaxLength >= intMaxCodeLen Then
                        ShowMsgbox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                        txtCode.Text = Space(txtCode.MaxLength)
                       chkChangeCode.Value = 0
                       chkChangeCode.Enabled = False
                    Else
                        ShowMsgbox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                        txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                        txtCode.MaxLength = txtCode.MaxLength + 1
                        txtCode.Tag = txtCode.MaxLength
                       chkChangeCode.Value = 1
                    End If
                Else
                    txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
                End If
            End If
        End With
        Exit Sub
   End If
   'ȷ���ϼ�����
   
    gstrSQL = "Select ���� From " & strTableName & " where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    
    If Not rsTemp.EOF Then
        txtUpCode.Text = zlCommFun.Nvl(rsTemp!����)
    End If
    
    '��ȷ���Ƿ����¼�
    gstrSQL = "select nvl(max(����),'') as ����  From " & strTableName & " Where  �ϼ�ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    intMaxCodeLen = rsTemp.Fields("����").DefinedSize

    If zlCommFun.Nvl(rsTemp!����) = "" Then
        '�������¼�
        '�����ϼ�IDȡ�ϼ�����
'        gstrSQL = "Select ���� From " & strTableName & " where id=" & lng�ϼ�id
'        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
'        txtUpCode.Text = zlCommFun.Nvl(rsTemp!����)
        txtCode.MaxLength = intMaxCodeLen - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If txtCode.MaxLength > 1 Then
            txtCode.Text = "01"
        Else
            txtCode.Text = "1"
        End If
        chkChangeCode.Value = 1
        chkChangeCode.Enabled = False
        Exit Sub
    End If
    
    With rsTemp
        txtCode.MaxLength = Len(!����) - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If Mid(!����, Len(txtUpCode.Text) + 1) = String(txtCode.MaxLength, "9") Then
            If Len(txtUpCode.Text) + txtCode.MaxLength >= intMaxCodeLen Then
                ShowMsgbox "�÷����¼�������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                txtCode.Text = Space(txtCode.MaxLength)
               chkChangeCode.Value = 0
               chkChangeCode.Enabled = False
            Else
                ShowMsgbox "�÷����¼��������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                txtCode.MaxLength = txtCode.MaxLength + 1
                txtCode.Tag = txtCode.MaxLength
               chkChangeCode.Value = 1
            End If
        Else
            txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ImeLanguage(ByVal blnOpen As Boolean)
    '-----------------------------------------------------------------------------------
    '����: ��/�ر����뷨
    '����: blnOpen-�Ǵ򿪻��ǹر�(trueΪ��,falseΪ�ر�)
    '���أ�
    '-----------------------------------------------------------------------------------
    If blnOpen Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '���ܣ����ı���ĵ��ı�ѡ�л����������뷨
    '����:blnOpenIme-�Ƿ�����뷨
    '����:
    '--------------------------------------------------------------------------------------------------------
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text) ' Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
    If blnOpenIme Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

'Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
'    '-----------------------------------------------------------------------------------
'    '����:ȡĳ�ֶε�ֵ
'    '����:rsObj          �������ֶ�
'    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
'    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
'    '-----------------------------------------------------------------------------------
'    If IsNull(rsObj) Then
'        Nvl = varValue
'    Else
'        Nvl = rsObj
'    End If
'End Function

'Public Function Dec2Bin(bDec As Byte) As String
'    '���ܣ�ʮ����תΪ�����ƺ���
'    '�÷���String  Dec2Bin(Bdec as Byte)
'    '���أ�  ʮ���ƵĶ����� �ַ���(String)
'    '����  ����"0"
'    Dim strBin As String
'
'    On Error GoTo Err
'    If bDec > 255 Then
'        Dec2Bin = "-1"
'        Exit Function
'    End If
'    strBin = ""
'    'תΪ�ַ���
'    While bDec > 0
'        strBin = bDec Mod 2 & strBin
'        bDec = Fix(bDec / 2)
'    Wend
'    '������8λ
'    If Len(strBin) < 9 Then
'        While Len(strBin) < 8
'            strBin = "0" & strBin
'        Wend
'    End If
'    Dec2Bin = strBin
'    Exit Function
'Err:
'   Dec2Bin = "0"
'End Function
'
'Public Function Bin2Dec(strBin As String) As Long
'    '���ܣ�������תΪʮ���ƺ���
'    '�÷���Long  bin2dec(strBin as String)
'    '���أ�  �����Ƶ�ʮ���� ��������Long��
'    '����  ����-1
'    Dim lDec As Long
'    Dim lCount As Long
'    Dim i As Long
'
'    On Error GoTo ErrHand
'    lDec = 0
'    If strBin = "" Then strBin = "0"
'    lCount = Len(strBin)
'    For i = 1 To lCount
'        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
'        strBin = Right(strBin, Len(strBin) - 1)
'        DoEvents
'    Next
'    Bin2Dec = lDec
'    Exit Function
'ErrHand:
'    Bin2Dec = -1
'End Function

Public Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer, Optional blnNum As Boolean = False)
    '----------------------------------------------------------------------------------------------------------------
    '������������ָ�����н�������
    '���������mshFilter-ָ��������
    '          intPreCol-�ϴ���
    '           intPreSort-�ϴ�����
    '           blnNum-�Ƿ�Ϊ������
    '���������
    '���أ�
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            strTemp = .TextMatrix(.Row, 0)
            If blnNum Then
                If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                   .Sort = flexSortNumericAscending
                   intPreSort = flexSortNumericAscending
                Else
                   .Sort = flexSortNumericDescending
                   intPreSort = flexSortNumericDescending
                End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            .Row = FindRow(mshFilter, strTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Public Function FindRow(ByVal mshgrd As MSHFlexGrid, ByVal varTemp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------------------------------------------
    '�������������ҷ�����������
    '���������varTemp-ָ����ֵ
    '           mshGrd-ָ������
    '           intCol-ָ������
    '���������
    '���أ��ɹ������ҵ�����
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intTmp As Integer
    
    With mshgrd
        For intTmp = 1 To .Rows - 1
            If IsDate(varTemp) Then
               If Format(.TextMatrix(intTmp, intCol), "yyyy-mm-dd") = Format(varTemp, "yyyy-mm-dd") Then
                  FindRow = intTmp
                  Exit Function
               End If
            Else
                If .TextMatrix(intTmp, intCol) = varTemp Then
                  FindRow = intTmp
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Public Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    Err = 0
    On Error GoTo ErrHand:
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    Exit Function
ErrHand:
    TranNumToDate = ""
End Function

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
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
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
    Err = 0
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
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function Check���Ȩ��(ByVal strPrivs As String, ByVal strPrv As String) As Boolean
    '����:���Ȩ���Ƿ����
    Dim strTmp As String
    strTmp = strPrv
    If IsNumeric(strPrv) Then
        '1λ--ҩƷ��Ӧ�̡���2λ--���ʹ�Ӧ�̡���3λ--�豸��Ӧ�̡���4λ--����,   5-�������� ÿλ��1�����ʾ,1��ʾΪtrue,0Ϊfalse,�Ժ�����ϵͳ�ӵ�6λ��ʼ
        strTmp = Decode(strPrv, 1, "ҩƷ", 2, "����", 3, "�豸", 4, "����", 5, "����", "��Ȩ��")
    End If
    Check���Ȩ�� = InStr(1, ";" & strPrivs & ";", ";" & strPrv & ";") <> 0
End Function

Public Function Get����Ȩ��(ByVal strPrivs As String, Optional aliasName As String = "", Optional bln��Ӧ�� As Boolean = True) As String
    '����:���Ȩ���Ƿ����
    Dim strTmp As String
    
    'Ӧ����¼�еĸ����ʶ:1����ҩƷӦ����   2��������Ӧ����   3�����豸Ӧ����   4��������,5--��������
    strTmp = ""
    If InStr(1, ";" & strPrivs & ";", ";ҩƷ;") <> 0 Then
        If bln��Ӧ�� Then
            strTmp = strTmp & " or substr(" & aliasName & "����,1,1)=1"
        Else
            strTmp = strTmp & " ,1"
        End If
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";����;") <> 0 Then
        If bln��Ӧ�� Then
            strTmp = strTmp & " or substr(" & aliasName & "����,2,1)=1"
        Else
            strTmp = strTmp & " ,2"
        End If
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";�豸;") <> 0 Then
        If bln��Ӧ�� Then
            strTmp = strTmp & " or substr(" & aliasName & "����,3,1)=1"
        Else
            strTmp = strTmp & " ,3"
        End If
        
    End If
    If InStr(1, ";" & strPrivs & ";", ";����;") <> 0 Then
        If bln��Ӧ�� Then
            strTmp = strTmp & " or substr(" & aliasName & "����,4,1)=1"
        Else
            strTmp = strTmp & " ,4"
        End If
        
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";����;") <> 0 Then
        If bln��Ӧ�� Then
            strTmp = strTmp & " or substr(" & aliasName & "����,5,1)=1"
        Else
            strTmp = strTmp & " ,5"
        End If
        
    End If
    If strTmp <> "" Then
        If bln��Ӧ�� Then
            strTmp = "  (" & Mid(strTmp, 4) & ") "
        Else
            strTmp = " NVL(" & aliasName & "ϵͳ��ʶ,4)  in (" & Mid(strTmp, 3) & ") "
        End If
    Else
        strTmp = " 1=2 "
    End If
    
    Get����Ȩ�� = strTmp
End Function

'Public Function Decode(ParamArray arrPar() As Variant) As Variant
''���ܣ�ģ��Oracle��Decode����
'    Dim varValue As Variant, i As Integer
'
'    i = 1
'    varValue = arrPar(0)
'    Do While i <= UBound(arrPar)
'        If i = UBound(arrPar) Then
'            Decode = arrPar(i): Exit Function
'        ElseIf varValue = arrPar(i) Then
'            Decode = arrPar(i + 1): Exit Function
'        Else
'            i = i + 2
'        End If
'    Loop
'End Function

Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "", Optional TxtAlignment As mAlignment = 1, Optional blnFontBold As Boolean = False)
    '���ܣ���PictureBoxģ���3Dƽ�水ť
    '������intStyle:0=ƽ��,-1=����,1=͹��,2-��͹��
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            Select Case IntStyle
            Case 1
                DrawEdge .hDC, PicRect, BDR_RAISEDINNER Or BF_SOFT, BF_RECT
            Case 2
                DrawEdge .hDC, PicRect, EDGE_RAISED, BF_RECT
            Case -1
                DrawEdge .hDC, PicRect, BDR_SUNKENOUTER Or BF_SOFT, BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) - 10
            End If
            .FontBold = blnFontBold
            picBox.Print strName
        End If
    End With
End Sub

Public Function Check������Ӧ����ϸ(ByVal lng������� As Long) As Boolean
    '-------------------------------------------------------------------------------------
    '����:��鸶����ϸ�ܽ����Ӧ����ϸ�ܽ��֮���Ƿ����!
    '����:lng�������-�������
    '����:���,����true�����򷵻�false
    '-------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim dbl������ As Double
    
    On Error GoTo errHandle
    strSQL = "Select sum(nvl(a.���,0)) AS ������ from �����¼ a where �������=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", lng�������)
    If rsTemp.EOF Then
        ShowMsgbox "��������صĸ��������(�������:" & lng������� & ")!"
        Exit Function
    End If
    dbl������ = Val(Nvl(rsTemp!������))
     
    strSQL = "Select Sum(Case When ��¼���� = 2 Then �ƻ���� " & _
             "                When not ��¼���� In (-1, 2) And Nvl(�ƻ����, 0) <> nvl(��Ʊ���,0) and �ƻ���� is null then ��Ʊ��� " & _
             "                When not ��¼���� In (-1, 2) And Nvl(�ƻ����, 0) <> nvl(��Ʊ���,0) and �ƻ���� is not null then �ƻ���� " & _
             "                Else 0 End) ��Ʊ��� " & _
             "From Ӧ����¼ " & _
             "Where ������� = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", lng�������)
    If rsTemp.EOF Then
        ShowMsgbox "��������ص�Ӧ����ϸ������(�������:" & lng������� & ")!"
        Exit Function
    End If
    
    If Round(dbl������, 2) <> Round(Val(Nvl(rsTemp!��Ʊ���)), 2) Then
        Call ShowMsgbox("���θ���(" & Format(dbl������, "###0.00;-###0.00;0;0") & ")�뱾�θ������ϸ�ܶ�(" & Format(Round(Val(Nvl(rsTemp!��Ʊ���)), 2), "####0.00;-###0.00;0;0") & ")���ȣ�����(�������:" & lng������� & ")!")
        Exit Function
    End If
    
    Check������Ӧ����ϸ = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

'Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
''���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
'    Dim vRect As RECT
'    Call GetWindowRect(lngHwnd, vRect)
'    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
'    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
'    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
'    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
'    GetControlRect = vRect
'End Function

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

''ȡ���ݿ��з�Ʊ�ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
'Public Function Get��Ʊ��Len() As Integer
'    Dim rsTemp As New Recordset
'
'    On Error GoTo errHandle
'    gstrSQL = "select ��Ʊ�� from Ӧ����¼ where rownum<1 "
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȡ�ֶγ���"
'    Get��Ʊ��Len = rsTemp.Fields(0).DefinedSize
'    rsTemp.Close
'    Exit Function
'
'errHandle:
'    If ErrCenter = 1 Then Resume
'End Function

Public Sub zlInitSystemPara()
    '------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ����
    '����:���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    With gSystemPara
        '0-ƴ����,1-�����,2-����
        .int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ"))
        .bln���Ի���� = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
        '��1λ1-ȫ����ֻ�����,��2λ1-ȫ��ĸֻ�����,��HIS��������������
        .Para_���뷽ʽ = zlDatabase.GetPara(44, glngSys, 0, "11")
        '.Para_���뷽ʽ = IIf(.Para_���뷽ʽ = "", "11", .Para_���뷽ʽ)
     End With
     
     '���绯վ����Ϣ
     Call Initվ����Ϣ
End Sub

Public Sub ��ʼС��λ��()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼС��λ��
    '���:
    '����:
    '����:
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_С��λ��
        .�ɱ���С�� = 7
        .���ۼ�С�� = 7
        .���С�� = 4
        .����С�� = 3
    End With
    With gVbFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, False)
        .FM_��� = GetFmtString(g_���, False)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, False)
        .FM_���� = GetFmtString(g_����, False)
    End With
    With gOraFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, True)
        .FM_��� = GetFmtString(g_���, True)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, True)
        .FM_���� = GetFmtString(g_����, True)
    End With
End Sub

Public Function GetFmtString(ByVal С������ As gС������, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '����:����ָ����С����ʽ��
    '���: lngС��λ��-С��λ��
    '     blnOracle-������oracle�ĸ�ʽ������Vb�ĸ�ʽ��
    '����:
    '����:����ָ���ĸ�ʽ��
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim intλ�� As Integer
    Select Case С������
    Case g_����
         intλ�� = g_С��λ��.����С��
    Case g_���
         intλ�� = g_С��λ��.���С��
    Case g_�ɱ���
         intλ�� = g_С��λ��.�ɱ���С��
    Case g_�ۼ�
         intλ�� = g_С��λ��.���ۼ�С��
    Case Else
        intλ�� = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(intλ��, "9") & "'"
    Else
       GetFmtString = "#0." & String(intλ��, "0") & ";-#0." & String(intλ��, "0") & "; ;"
    End If
End Function

'Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
'    '------------------------------------------------------------------------------
'    '����:�жϿؼ��Ƿ��
'    '����:����ɹ�,����true,���򷵻�False
'    '����:���˺�
'    '����:2008/01/24
'    '------------------------------------------------------------------------------
'    Dim rsTemp As New ADODB.Recordset
'    Err = 0: On Error GoTo ErrHand:
'    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'End Function

'Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
'    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
'    Err = 0: On Error Resume Next
'    If blnDoEvnts Then DoEvents
'    If zlControl.IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
'End Sub

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Sub Initվ����Ϣ()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��վ��������Ϣ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-09-01 11:32:00
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gSystemPara.bln����վ�� = gstrNodeNo <> "-"
 End Sub
 
Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function

Public Function zlSelectDept(ByVal FrmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(GetMatchingSting(strSearch, False), "%", "*")
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(FrmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

Public Function GetStoreInfo(ByVal strClass As String) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select Distinct a.Id, a.����, a.����, a.���� " & _
             "From ���ű� A, ��������˵�� B, �������ʷ��� C " & _
             "Where a.Id = b.����id And c.���� = b.�������� And c.���� In (" & strClass & ") " & zl_��ȡվ������(True, "A") & " " & _
             "Order By a.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���пⷿ��Ϣ")
    
   Set GetStoreInfo = rsTmp.Clone
    
    Exit Function

errHandle:
    If ErrCenter = 1 Then Resume
End Function
