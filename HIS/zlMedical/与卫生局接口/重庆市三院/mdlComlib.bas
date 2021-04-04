Attribute VB_Name = "mdlComlib"

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "���ַ���", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Sub TxtSelAll(objTxt As Object)
'���ܣ����༭��ĵ��ı�ȫ��ѡ��
'������objTxt=��Ҫȫѡ�ı༭�ؼ�,�ÿؼ�����SelStart,SelLength����
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    rs.Open "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", gcnOracle
    GetMaxLength = rs.Fields(0).DefinedSize
    
End Function

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = objVsf.BackColorSel
    End If
    
End Sub

Public Function GetNextId(strTable As String) As Long
    '------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
    '������
    '   strTable��������
    '���أ�
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & Trim(strTable) & "_ID.Nextval From Dual"
    
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    
    GetNextId = rsTmp.Fields(0).Value
    Exit Function
errH:
    
End Function

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
    
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
'��ʾ��������
'ChmName:CHM��ʽ�ļ�
'SHwnd:���봰�ھ��(��Ϊ��������)
'htmName:��ӳ��CHM�е�htm�ļ�����

    Dim Path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zlPiesFlat" & ".chm"
    If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, Path, &H0, htmName & ".htm")
    
    ShowHelp = True
    Exit Function

ShowHelpErr:
    Err.Clear
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ����洰�弰���и��ֿؼ���״̬
'������objForm:Ҫ����Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    
    Dim objThis As Object
    Dim strTmp As String
    Dim strIndex As String
    Dim i As Integer, strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    
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
    Dim strIndex As String
    Dim strSQL As String
    Dim strOEM As String
    
    On Error Resume Next
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '�ָ������״̬��λ�á���С
    strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    
    aryInfo = Split(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.Name & strUserDef & "\Form", "״̬", strTmp), ",")
    
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

    RestoreWinState = True
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPrivs As String
    Dim strWhere As String
    
    On Error GoTo errH
    
    'strWhere = zlRegFunctions(GetUnitInfo("ע����"))
    
    strWhere = "1=1"
    
    If strWhere = "" Or strWhere = "-" Then Exit Function
    
        strSQL = _
            "Select Distinct ���� From (" & _
            " Select A.ϵͳ,A.���,A.����" & _
            " From zlRoleGrant A,Session_Roles B" & _
            " Where A.��ɫ = B.Role And A.���=" & lngProgID & " And A.ϵͳ=" & lngSys & _
            " Union All" & _
            " Select A.ϵͳ,B.���,B.����" & _
            " From zlPrograms A,zlProgFuncs B" & _
            " Where A.���=B.��� And A.ϵͳ=B.ϵͳ And A.���=" & lngProgID & " And A.ϵͳ=" & lngSys & _
            " And (Exists(Select 1 From Session_Roles Where Role='DBA')" & _
            " Or A.ϵͳ in (Select ��� From zlSystems Where Upper(������)=USER)" & _
            ")) Where " & strWhere
    
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle
    Do While Not rsTmp.EOF
        strPrivs = strPrivs & ";" & rsTmp!����
        rsTmp.MoveNext
    Loop
    GetPrivFunc = Mid(strPrivs, 2)
    Exit Function
errH:
    
    ShowSimpleMsg Err.Description
    
End Function
