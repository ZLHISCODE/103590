Attribute VB_Name = "mdlPublic"
Option Explicit


Public Const SW_RESTORE = 9
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const SW_HIDE = 0


Public Const EXT_LIKEWAY = "ģ��ƥ�䷽ʽ"
Public Const EXT_PRO_VALUE_LEFTWAY = "��ƥ��"
Public Const EXT_PRO_VALUE_RIGHTWAY = "��ƥ��"
Public Const EXT_PRO_VALUE_FULLWAY = "�м�ƥ��"

Public Const EXT_DATERANGE = "���ڷ�Χ�޶�"
Public Const EXT_UPPERCONVERT = "��дת��"
Public Const EXT_NUMBERCONVERT = "����ת��"
Public Const EXT_IGNORESYSPAR = "����ϵͳ����"

Public gstrPara As String              'ϵͳ����
Public gstrBasePara As String          'ҵ�����
Public gstrCachePath As String         '�洢�����Ļ���·��
Public gbytFontSize As Byte

Public glngUserId As Long
Public gstrUserAccount As String
Public gstrUserName As String

Public glngSysNo As Long
Public glngModuleNo As Long
Public gstrDeptId As String     'Ӱ��ҽ��ѡ��ȫ������ʱ������Ҫ��¼������ҵ�ID

Public gblnTimeChanged As Boolean

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long

Public Const G_LNG_PACSSTATION_MODULE As Long = 1290    'Ӱ��ҽ��ϵͳ���
Public Const G_LNG_VIDEOSTATION_MODULE As Long = 1291   'Ӱ��ɼ�ϵͳ���
Public Const G_LNG_PATHSTATION_MODULE As Long = 1294    'Ӱ����ϵͳ���
'PACS�Զ����ѯ���Ƚ϶����Ĳ˵���
Public Const conMenu_PacsQuery_TimeLab = 8243     'ʱ���ǩ
Public Const conMenu_PacsQuery_TimeCbo = 8244     'ʱ��������
Public Const conMenu_PacsQuery_TimeCustom = 8245     '�Զ���ʱ������
Public Const conMenu_PacsQuery_FindWay = 8246          '��ѯ��ʽ
Public Const conMenu_PacsQuery_PatiControl = 8247          'Pati�ؼ�
Public Const conMenu_PacsQuery_Do = 8248          'ִ�в�ѯ
Public Const CB_GETCURSEL = &H147
Public Const CB_SETCURSEL = &H14E
Public Const LB_SETCURSEL = &H186
Public Const LB_GETCURSEL = &H188

'��ѯ��������
Public Const varName_���ݿ����� As String = "���ݿ�����"   '���ݿ�����
Public Const varName_���ݿ��û��� As String = "���ݿ��û���"   '���ݿ��û���
Public Const varName_ģ��� As String = "ģ���"  'ģ���
Public Const varName_�û�ID As String = "�û�ID"    '�û�ID
Public Const varName_����ID As String = "����ID"         '����ID
Public Const varName_��ѯ����ID As String = "��ѯ����ID"         '��ѯ����ID
Public Const varName_��ѯ�������� As String = "��ѯ��������"         '��ѯ��������
Public Const varName_�б�ؼ��� As String = "�б�ؼ���"         '�б�ؼ���
Public Const varName_ϵͳ�� As String = "ϵͳ��"         'ϵͳ��
Public Const varName_�ֺ� As String = "�ֺ�"       '�ֺ�
Public Const varName_������ As String = "������"       '������
Public Const varName_�Ƿ����ù������� As String = "�Ƿ����ù�������"    '�Ƿ����ù�������


Public Const FE_FONTSMOOTHINGCLEARTYPE = 2
Public Const SPI_GETFONTSMOOTHINGTYPE = &H200A
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


Public Enum TNeedType
    tNeedName = 0
    tNeedNo = 1
    tNeedAll = 2
End Enum

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function SqlVerify(strSql As String, Optional ByVal blnVerifyHaveAdviceID As Boolean = False) As String
'blnVerifyHaveAdviceID �Ƿ񸽼���Ϣ��֤������ǣ�����[ϵͳ.ҽ��ID����֤]
    Dim objSqlParse As New clsSqlParse
    Dim i As Integer
    Dim strPars As String
    Dim strPar As String
    Dim rsRecord As Recordset
    
    strPars = ""
    On Error GoTo errRollback
        If Len(Trim(strSql)) = 0 Then
        SqlVerify = "û�в�ѯ���"
        Exit Function
    End If
    
    If blnVerifyHaveAdviceID Then
        If InStr(1, strSql, "ϵͳ.ҽ��ID") = 0 Then
            SqlVerify = "������Ϣ����������û��[ϵͳ.ҽ��ID]�����顣"
            Exit Function
        Else
            If InStr(1, strSql, "[ϵͳ.ҽ��ID]") = 0 Then
                SqlVerify = "������Ϣ����������[ϵͳ.ҽ��ID]����֮�䲻Ҫ�пո���������ַ������顣"
                Exit Function
            End If
        End If
    End If
    
    strSql = objSqlParse.GetTestSql(strSql)
    Set rsRecord = ExecuteSql(strSql, "��ѯ��֤")
    
    '��֤�ظ��ֶο�ʼ
    If rsRecord Is Nothing Then
        SqlVerify = "SQLû�в�ѯ����Ч����"
        Exit Function
    End If

    For i = 1 To rsRecord.Fields.Count
        strPar = rsRecord.Fields(i - 1).Name
        If InStr(";" & strPars & ";", ";" & strPar & ";") = 0 Then
            strPars = strPars & ";" & strPar
        Else
            SqlVerify = "SQL�����ظ��ֶ�[" & strPar & "]"
            Exit Function
        End If
    Next
    
    SqlVerify = ""
Exit Function
errRollback:
    SqlVerify = Err.Description
    SqlVerify = Mid(SqlVerify, InStr(1, SqlVerify, ":") + 2)
End Function

Public Function IsHaveID(strSql As String) As String
'��֤Sql���Ƿ����ҽ��ID�Ͳ���ID
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As Recordset
    Dim strItem As String
    Dim strPar As String
    Dim i As Long
    
    IsHaveID = ""
    
    objSqlParse.init strSql
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    
    For i = 0 To rsRecord.Fields.Count - 1
        strItem = strItem & ",[" & rsRecord.Fields(i).Name & "]"
    Next
    strItem = Mid(strItem, 2)
    If InStr(UCase(strItem), "[ҽ��ID]") = 0 Then
        IsHaveID = "��ѯ������뺬�С�ҽ��ID��"
        Exit Function
    End If

    If InStr(UCase(strItem), "[����ID]") = 0 Then
        IsHaveID = "��ѯ������뺬�С�����ID��"
        Exit Function
    End If
    
    For i = 1 To objSqlParse.SqlStruct.ParCount
        strPar = strPar & ",[" & objSqlParse.SqlStruct.AllParameter(i) & "]"
    Next
    strPar = Mid(strPar, 2)
    
    If InStr(UCase(strPar), "[ҽ��ID]") = 0 And InStr(UCase(strPar), "[ϵͳ.ҽ��ID]") = 0 Then
        IsHaveID = "��ѯ�������뺬�С�ҽ��ID��"
        Exit Function
    End If

    If InStr(UCase(strPar), "[����ID]") = 0 And InStr(UCase(strPar), "[ϵͳ.����ID]") = 0 Then
        IsHaveID = "��ѯ�������뺬�С�����ID��"
        Exit Function
    End If
End Function


Public Function CurServerDate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    CurServerDate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    CurServerDate = 0
    Err = 0
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function StationName() As String
'�������������
    Dim strComputer As String * 256
    
    Call GetComputerName(strComputer, 255)
    
    StationName = Replace(strComputer, Chr(0), "")
End Function


Public Function RunScripting(sctExecute As ScriptControl, ByVal strScript As String) As Variant
'ִ��vbs�ű�
    Dim strFormatScript As String
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
    
    strFormatScript = Trim$(strScript)
    
    lngStartIndex = InStr(strFormatScript, "[")
    lngEndIndex = InStr(strFormatScript, "]")
    
    If lngStartIndex <= 0 Or lngEndIndex <= 0 Then
        RunScripting = strScript
        Exit Function
    End If
    
    If lngStartIndex <> 1 And lngEndIndex <> Len(strFormatScript) Then
        RunScripting = strScript
        Exit Function
    End If
    
    RunScripting = Null
    
    strFormatScript = Replace(Replace(strFormatScript, "[", ""), "]", "")
    sctExecute.Reset
    
On Error GoTo errHandle
    RunScripting = sctExecute.Eval(strFormatScript)
    Exit Function
errHandle:
    strFormatScript = "function return()" & vbCrLf & strFormatScript & " end function"
    Call sctExecute.AddCode(strFormatScript)
    
    RunScripting = RunFunction(sctExecute)
End Function

Private Function RunFunction(sctExecute As ScriptControl) As Variant
On Error GoTo errHandle
    RunFunction = Null
    
    RunFunction = sctExecute.Run("return")
Exit Function
errHandle:
    RunFunction = Null
End Function


Public Function GetMinValue(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long, ByVal lngIndex3 As Long) As Long
'��ȡ��Сֵ
    Dim lngResult As Long
    Dim lngV1 As Long
    Dim lngV2 As Long
    Dim lngV3 As Long

    lngResult = 100000

    lngV1 = IIf(lngIndex1 > 0, lngIndex1, 100000)
    lngV2 = IIf(lngIndex2 > 0, lngIndex2, 100000)
    lngV3 = IIf(lngIndex3 > 0, lngIndex3, 100000)

    If lngResult > lngV1 Then lngResult = lngV1
    If lngResult > lngV2 Then lngResult = lngV2
    If lngResult > lngV3 Then lngResult = lngV3

    If lngResult = 100000 Then lngResult = 0

    GetMinValue = lngResult
End Function


Public Sub CopyStrArray(ByRef arySource() As String, ByRef aryTag() As String, Optional ByVal lngStart As Long = 1)
'��������
    Dim lngSourceCount As Long
    Dim lngUbound As Long
    Dim i As Long
    
    lngSourceCount = UBound(arySource)
    
    For i = lngStart To lngSourceCount
        lngUbound = UBound(aryTag) + 1
        ReDim Preserve aryTag(lngUbound)
        
        aryTag(lngUbound) = arySource(i)
    Next i
    
End Sub


Public Sub CopyLngArray(ByRef arySource() As Long, ByRef aryTag() As Long, Optional ByVal lngStart As Long = 1)
'��������
    Dim lngSourceCount As Long
    Dim lngUbound As Long
    Dim i As Long
    
    lngSourceCount = UBound(arySource)
    
    For i = lngStart To lngSourceCount
        lngUbound = UBound(aryTag) + 1
        ReDim Preserve aryTag(lngUbound)
        
        aryTag(lngUbound) = arySource(i)
    Next i
    
End Sub


Public Sub CopyVariantArray(ByRef arySource() As Variant, ByRef aryTarget() As Variant, _
                            Optional ByVal lngStart As Long = 1, Optional ByVal blnIsFirst As Boolean = False)
'���Ʊ�������
    Dim lngSourceCount As Long
    Dim lngUbound As Long
    Dim i As Long
    
    For i = lngStart To lngSourceCount
        If blnIsFirst Then
            aryTarget(lngUbound) = arySource(i)
        End If
        
        lngUbound = UBound(aryTarget) + 1
        ReDim Preserve aryTarget(lngUbound)
        
        If Not blnIsFirst Then
            aryTarget(lngUbound) = arySource(i)
        End If
    Next i
End Sub

Public Function IsUseClearType() As Boolean
    Dim lngCurType As Long

    Call SystemParametersInfo(SPI_GETFONTSMOOTHINGTYPE, 0, lngCurType, 0)
    IsUseClearType = IIf(lngCurType = FE_FONTSMOOTHINGCLEARTYPE, True, False)
   
End Function

Public Sub SeekIndexSimple(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
    Dim i As Long

    For i = 1 To objCbo.ListCount
        If objCbo.List(i) = strText Then
            objCbo.ListIndex = i
            Exit Sub
        End If
    Next
    
End Sub

Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional intIsSearchNo As TNeedType = tNeedName)
'���ܣ���ComboBox�в��Ҳ���λ
'������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
      'blnPreserve--����Ҳ���ƥ����Ŀ���򱣳�ԭ����Ŀ
      'intIsSearchNo -- 0:ͨ�����붨λ,1:ͨ�����ֶ�λ,2:�ù���������ֶ�λ
'˵����δ�ܶ�λʱ,����ListIndex=-1
'       Cbo.SeekIndex���ܱȽϼ򵥣�����index��ᴥ���¼������ʺ�ʹ��
    Dim i As Long

    For i = 0 To objCbo.ListCount - 1
        If IIf(Abs(intIsSearchNo) = tNeedAll, objCbo.List(i), IIf(Abs(intIsSearchNo) = tNeedNo, zlStr.NeedCode(objCbo.List(i)), zlStr.NeedName(objCbo.List(i)))) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    
    If blnPreserve = True Then
        If blnEvent = False Then
            Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.ListIndex)
        End If
    Else
        If blnEvent Then
            objCbo.ListIndex = -1
        Else
            Call zlControl.CboSetIndex(objCbo.hWnd, -1)
        End If
    End If
    
End Sub


Public Function GetExtPropertyValue(ByVal strExtProperty As String, ByVal strPropertyName As String) As String
    Dim i As Long
    Dim strPropertys() As String
    Dim strSplit() As String
    
On Error GoTo errHandle
    GetExtPropertyValue = ""
    
    If Trim(strExtProperty) = "" Then Exit Function
    
    strPropertys = Split(strExtProperty, ";")
    
    For i = 0 To UBound(strPropertys)
        If strPropertys(i) <> "" Then
            strSplit = Split(strPropertys(i), "=")
            If strSplit(0) = strPropertyName Then
                GetExtPropertyValue = Trim(strSplit(1))
                Exit Function
            End If
        End If
    Next
    
Exit Function
errHandle:
    GetExtPropertyValue = ""
End Function

Public Function SetListIndex(lst As control, ByVal NewIndex As Long) As Long
    If TypeOf lst Is ListBox Then
    
        Call SendMessage(lst.hWnd, LB_SETCURSEL, NewIndex, 0&)
        
        SetListIndex = SendMessage(lst.hWnd, LB_GETCURSEL, NewIndex, 0&)
        
        ElseIf TypeOf lst Is ComboBox Then
        
        Call SendMessage(lst.hWnd, CB_SETCURSEL, NewIndex, 0&)
        SetListIndex = SendMessage(lst.hWnd, CB_GETCURSEL, NewIndex, 0&)
    
    End If
End Function


































