Attribute VB_Name = "mdlSQLSource"

Option Explicit
Public Enum gEditType
     g���� = 0
     g�޸� = 1
     g�鿴 = 4
End Enum

'------------------------------------------------------------------------------------
Public grsObject As ADODB.Recordset '��ǰ�û�������SelectȨ�޵Ķ���(�����򵼻򷢲�)
'------------------------------------------------------------------------------------

Public gblnRunLog As Boolean '�Ƿ��¼ʹ����־
Public gblnErrLog As Boolean '�Ƿ��¼���д���


Public glngKeyHook As Long


Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����TAB���ĺ���
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE = 3


Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ

Public gblnOK As Boolean
Public Type CustomPar
    ���� As String
    ֵ�б� As String
    ����SQL As String
    ��ϸSQL As String
    �����ֶ� As String
    ��ϸ�ֶ� As String
    ���� As String
    ��ʽ As Byte
End Type
Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public Const GSTR_SBC = "���������������������������������������������£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣ڣ����������������������������������������������"
Public Const GSTR_DBC = "(+-*/=<>)!:1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcedfghijklmnopqrstuvwxyz;,.?|%#"

Public Function TLen(str As String) As Long
'���ܣ������ַ�������ʵ����
    TLen = LenB(StrConv(str, vbFromUnicode))
End Function

Public Function RemoveNote(ByVal strSQL As String) As String
'���ܣ��Ƴ�SQL����е�ע��
'˵����ֻ֧���Ƴ����е�ע��
    Dim strTmp As String, i As Integer
    Dim arrLine() As String
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbLf, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr, vbCrLf)
    arrLine = Split(strSQL, vbCrLf)
    
    For i = 0 To UBound(arrLine)
        If Not Trim(arrLine(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & arrLine(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function


Public Function TrimChar(str As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function


Public Function CheckPars(strSQL As String) As Boolean
'���ܣ����SQL����в�����"[]"�Ƿ����,�Լ��������Ƿ���ȷ(������,������)
    Dim intLeft As Integer, intRight As Integer
    Dim intMin As Integer, intMax As Integer
    Dim strTmp As String, strPar As String, strPars As String
    Dim i As Long
    
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "[" Then intLeft = intLeft + 1
        If Mid(strSQL, i, 1) = "]" Then intRight = intRight + 1
    Next
    
    If intLeft <> intRight Then Exit Function '"["��"]"�����
    
    If intLeft = 0 And intRight = 0 Then CheckPars = True: Exit Function
    
    strTmp = strSQL
    intMin = 32767
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(strPar) = "" Then
            strPar = 0
        ElseIf Not IsNumeric(strPar) Then
            Exit Function '�����ֱ��
        End If
        If CInt(strPar) < intMin Then intMin = CInt(strPar)
        If CInt(strPar) > intMax Then intMax = CInt(strPar)
        If InStr(strPars, "," & CInt(strPar)) = 0 Then strPars = strPars & "," & CInt(strPar)
    Loop
    If intMin <> 0 Then Exit Function '���Ǵ�0��ʼ���
    If strPars <> "" Then strPars = Mid(strPars, 2)
    If UBound(Split(strPars, ",")) <> intMax Then Exit Function '�����������
    CheckPars = True
End Function


Public Function SQLObject(ByVal strSQL As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    
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



Public Function CheckObjectPriv(strObject As String) As String
'���ܣ���鵱ǰ�û���ָ�������Ƿ���ȫ��Ȩ�޷���
'������strObject=��������,��"���ű�,���˷��ü�¼"
'���أ���ȫ=��,����ȫ=���ܷ��ʵĶ�����,��"���ű�,���˷��ü�¼"
'˵����������У������Դ֮ǰ����Ƿ���Ȩ�޲�ѯSQL����еĶ���
'�ο���grsObject
    Dim i As Integer
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") = 0 Then
                grsObject.Filter = "OBJECT_NAME='" & Split(strObject, ",")(i) & "'"
            Else
                '�������ͼ���������ǰ׺,����������߶���Ȩ��
                grsObject.Filter = "OWNER='" & Split(Split(strObject, ",")(i), ".")(0) & _
                    "' And OBJECT_NAME='" & Split(Split(strObject, ",")(i), ".")(1) & "'"
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

Public Function ObjectOwner(strObject As String, Optional frmParent As Object) As String
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
                grsObject.Filter = "OBJECT_NAME='" & Split(strObject, ",")(i) & "'"
                If grsObject.RecordCount = 1 Then
                    If InStr(ObjectOwner & ",", "," & grsObject!owner & "." & Split(strObject, ",")(i) & ",") = 0 Then
                        ObjectOwner = ObjectOwner & "," & grsObject!owner & "." & Split(strObject, ",")(i)
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

Public Function SQLReplaceOwner(ByVal strSQL As String, strOwner As String) As String
'���ܣ���SQL����滻�ɴ����������ߵ���ʽ
'������strSQL=ԭʼSQL���,strOwner=���������ߴ�,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼"
'���أ����ʶ������������ǰ׺��SQL���
'˵����1.����������ֱ��ִ���û�SQL���,������Ҫ��Ȩ�����˽��ͬ��ʡ�
'      2.�Ա������ֶ�����ͬ���ֶ���û�д������,������
    Dim i As Long, j As Long
    Dim intLoc As Long, blnDo As Boolean
    
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
    SQLReplaceOwner = strSQL
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


Public Sub CopyPars(ByVal objSPars As RPTPars, ByRef objOPars As RPTPars)
'���ܣ���������������
    Dim tmpPar As RPTPar
    
    Set objOPars = New RPTPars
    For Each tmpPar In objSPars
        With tmpPar
            objOPars.Add .����, .���, .����, .����, .ȱʡֵ, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ�, .��ϸ�ֶ�, .����, "_" & .Key, .Reserve
        End With
    Next
End Sub


Public Function GetExecSQL(ByVal strSQL As String, Optional ByVal objPars As RPTPars) As String
'���ܣ���ȡ��ִ�е�SQL
    Dim rstmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim strPar As String, bytPar As Byte, i As Integer
    
    strCheck = strSQL
    On Error GoTo errH
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            strPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(strPar) = "" Then strPar = 0
            bytPar = CByte(strPar)
            
            '��ȱʡ����ֵ�滻
            If objPars("_" & CInt(bytPar)).ȱʡֵ <> "" And Not objPars("_" & CInt(bytPar)).ȱʡֵ Like "*��" Then
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        strPar = "'" & Replace(objPars("_" & CInt(bytPar)).ȱʡֵ, "'", "''") & "'"
                    Case 1 '����
                        strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                    Case 2 '����
                        If Left(objPars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                            strPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).ȱʡֵ)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).ȱʡֵ, ":") > 0 Then
                                '��ʱ���ʽ
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '��ʱ���ʽ
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '������
                        strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                End Select
            Else 'ȱʡֵΪ�ջ�Ϊ�Զ�����
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        strPar = "'�մ�'"
                    Case 1 '����
                        strPar = 1 '����Ϊ0���ܵ��³���Ϊ0
                    Case 2 '����
                        strPar = "Sysdate"
                    Case 3 '������(ֱ���滻)
                        If objPars("_" & CInt(bytPar)).ȱʡֵ = "�̶�ֵ�б�" Then
                            'ȡ�̶�ֵ�е�ȱʡֵ
                            '���õķָ���
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), 1) = "��" Then
                                    strPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            'û������ȱʡֵ��ȡ��һ��
                            If strPar = "" Then
                                strPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).ȱʡֵ = "ѡ�������塭" Then
                            If objPars("_" & CInt(bytPar)).ֵ�б� <> "" Then
                                'ȡȱʡ��ֵ
                                strPar = Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).��ϸSQL <> "" And objPars("_" & CInt(bytPar)).��ϸ�ֶ� <> "" Then
                                strPar = GetDefaultValue(objPars("_" & CInt(bytPar)).��ϸSQL, objPars("_" & CInt(bytPar)).��ϸ�ֶ�)
                                If strPar <> "" Then strPar = CStr(Split(strPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).��ʽ = 1 Then
                                    strPar = " In (" & strPar & ") "
                                End If
                            Else
                                strPar = ""
                            End If
                        Else
                            strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                        End If
                End Select
            End If
            strCheck = strLeft & strPar & strRight
        Loop
    End If
    
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    GetExecSQL = strCheck
    Exit Function
errH:
    Err.Clear
    GetExecSQL = ""
End Function

Public Function CheckSQL(ByVal strSQL As String, strErr As String, Optional ByVal objPars As RPTPars) As String
'���ܣ�����ȱʡ�������SQL�����д�Ƿ���ȷ
'���أ�
'     �ɹ�=SQL���ֶδ�,�����˸����ֶε����Ƽ�����,��ʽ��"����,111|����,111|����,123",����ֵ��ADO.Field.TypeΪ׼
'     ʧ��=��
    Dim rstmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim strPar As String, bytPar As Byte, i As Integer
    
    strCheck = strSQL
    
    On Error GoTo errH
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            strPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(strPar) = "" Then strPar = 0
            bytPar = CByte(strPar)
            
            '��ȱʡ����ֵ�滻
            If objPars("_" & CInt(bytPar)).ȱʡֵ <> "" And Not objPars("_" & CInt(bytPar)).ȱʡֵ Like "*��" Then
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        strPar = "'" & Replace(objPars("_" & CInt(bytPar)).ȱʡֵ, "'", "''") & "'"
                    Case 1 '����
                        strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                    Case 2 '����
                        If Left(objPars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                            strPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).ȱʡֵ)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).ȱʡֵ, ":") > 0 Then
                                '��ʱ���ʽ
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '��ʱ���ʽ
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '������
                        strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                End Select
            Else 'ȱʡֵΪ�ջ�Ϊ�Զ�����
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        strPar = "'�մ�'"
                    Case 1 '����
                        strPar = 1 '����Ϊ0���ܵ��³���Ϊ0
                    Case 2 '����
                        strPar = "Sysdate"
                    Case 3 '������(ֱ���滻)
                        If objPars("_" & CInt(bytPar)).ȱʡֵ = "�̶�ֵ�б�" Then
                            'ȡ�̶�ֵ�е�ȱʡֵ
                            '���õķָ���
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), 1) = "��" Then
                                    strPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            'û������ȱʡֵ��ȡ��һ��
                            If strPar = "" Then
                                strPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).ȱʡֵ = "ѡ�������塭" Then
                            If objPars("_" & CInt(bytPar)).ֵ�б� <> "" Then
                                'ȡȱʡ��ֵ
                                strPar = Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).��ϸSQL <> "" And objPars("_" & CInt(bytPar)).��ϸ�ֶ� <> "" Then
                                strPar = GetDefaultValue(objPars("_" & CInt(bytPar)).��ϸSQL, objPars("_" & CInt(bytPar)).��ϸ�ֶ�)
                                If strPar <> "" Then strPar = CStr(Split(strPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).��ʽ = 1 Then
                                    strPar = " In (" & strPar & ") "
                                End If
                            Else
                                strPar = ""
                            End If
                        Else
                            strPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                        End If
                End Select
            End If
            strCheck = strLeft & strPar & strRight
        Loop
    End If
    
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next
    Call zlDatabase.OpenRecordset(rstmp, strCheck, "mdlPublic_CheckSQL") '�滻�ɵĶ��ǹ̶�����,ͬһ����Դһ�㲻��,����SQLҲ�����������
    strSQL = strCheck   '���ؿ�ִ��SQL
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rstmp.Fields
            If InStr(tmpFld.Name, "|") > 0 Then
                strErr = "�ֶ�""" & tmpFld.Name & """û�б�����"
                CheckSQL = "": Exit Function
            ElseIf InStr(tmpFld.Name, "'") > 0 Or InStr(tmpFld.Name, """") > 0 Then
                strErr = "�ֶ��� " & tmpFld.Name & " �Ƿ���"
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
    Exit Function
errH:
    Err.Clear
    strErr = "��������Դ�Ĳ����������󣬿�����SQL�еĲ��������ݱ��в����ڡ�"
    CheckSQL = ""
End Function


Public Function GetParSQL(ByVal strSQL As String) As String
'���ܣ���SQL���ɴ������ĸ�ʽ
'Select * FRom ���ű� Where ID=/*B1*/413/*E1*/
'Select * FRom ���ű� Where ID=[1]
    Dim strTmp As String, i As Integer
    Dim strL As String, strR As String
    Dim intMax As Integer
    
    On Error Resume Next
    
    strTmp = strSQL: intMax = -1
    Do While InStr(strTmp, "/*B") > 0
        strL = Left(strTmp, InStr(strTmp, "/*B") - 1)
        strR = Mid(strTmp, InStr(strTmp, "/*B") + 3)
        If Val(strR) > intMax Then intMax = Val(strR)
        strTmp = strL & strR
    Loop
    
    For i = 0 To intMax
        Do While InStr(strSQL, "/*B" & i & "*/") > 0
            strL = Left(strSQL, InStr(strSQL, "/*B" & i & "*/") - 1)
            strR = Mid(strSQL, InStr(strSQL, "/*E" & i & "*/") + Len("/*E" & i & "*/"))
            strSQL = strL & "[" & i & "]" & strR
        Loop
    Next
    
    GetParSQL = strSQL
End Function

Public Function InString(strText As String, strChars As String) As Boolean
'���ܣ������strText���Ƿ����strChars��ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strChars)
        If InStr(strText, Mid(strChars, i, 1)) > 0 Then
            InString = True
            Exit Function
        End If
    Next
End Function

Public Function GetDBVer() As Long
    Dim strSQL As String, rstmp As ADODB.Recordset
    
    strSQL = "Select To_Number(Replace(Substr(Banner, 6, 4), '.', '')) As Dbver From V$version Where Substr(Banner, 1, 4) = 'CORE'"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    GetDBVer = Val(rstmp!dbver)
    '10G:102,9i:92
    Exit Function
errH:
    Err.Clear
    GetDBVer = 102
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, msg, wp, lp)
End Function


Public Function GetParSQLMacro(str As String) As String
'����:�������������,������ת�������SQL����п��õ�ֵ
    Dim curDate As Date
    
    If InStr(str, "&") = 0 Then GetParSQLMacro = str: Exit Function
    
    curDate = Currentdate
    
    Select Case str
        Case "&��ǰ����"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��ǰ����ʱ��"
            GetParSQLMacro = "Sysdate"
        Case "&���쿪ʼʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�������ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ�쿪ʼʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ��ͬʱ��"
            GetParSQLMacro = "Sysdate-1"
        Case "&��һ��ͬʱ��"
            GetParSQLMacro = "Sysdate+1"
        Case "&��һ�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate + 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&��һ������"
            GetParSQLMacro = "Trunc(Sysdate+1)"
        Case "&ǰһ������"
            GetParSQLMacro = "Trunc(Sysdate - 7)"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "Trunc(Sysdate + 7)"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&���³�ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function
    

Public Function GetDefaultValue(ByVal strSQL As String, ByVal strFld As String, Optional ByVal strDefBand As String) As String
'���ܣ����ݲ���ѡ����SQL���壬������ʾ�ֶμ����ֶε�ֵ
'������strFld=��������Դ�ֶ�˵����
'      strDefBand=�������ȱʡ��ֵ,�Ƿ񰴴�ֵ����
'���أ���ʾֵ|��ֵ|ԭʼ��¼��
    Dim rstmp As New ADODB.Recordset
    Dim strTmp As String, i As Long
    Dim strShow As String, strBand As String
        
    'ȡ����ʾ,���ֶ���
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then strShow = CStr(Split(strTmp, ",")(0))
        If Split(strTmp, ",")(2) Like "*&B*" Then strBand = CStr(Split(strTmp, ",")(0))
    Next
    If strShow = "" And strBand = "" Then Exit Function
        
    '�򿪲�������Դ
    On Error GoTo errH
    strSQL = Replace(RemoveNote(strSQL), "[*]", "")
    Call zlDatabase.OpenRecordset(rstmp, strSQL, "mdlPublic_GetDefaultValue")  '[*]��SQL��''��,�����޷�����
    i = rstmp.RecordCount 'ԭʼ��¼����
        
    '�Ȱ�ָ���İ�ֵ���˳�������
    If Not rstmp.EOF And strDefBand <> "" Then
        If IsType(rstmp.Fields(strBand).Type, adVarChar) Then
            rstmp.Filter = strBand & "='" & strDefBand & "'"
        ElseIf IsType(rstmp.Fields(strBand).Type, adNumeric) Then
            If Not IsNumeric(strDefBand) Then Exit Function
            rstmp.Filter = strBand & "=" & strDefBand
        ElseIf IsType(rstmp.Fields(strBand).Type, adDBTimeStamp) Then
            If Not IsDate(strDefBand) Then Exit Function
            rstmp.Filter = strBand & "=#" & strDefBand & "#"
        End If
    End If
    
    '�ٷ���ȱʡ�����ݻ����������
    If Not rstmp.EOF Then
        strShow = Nvl(rstmp.Fields(strShow).Value, "")
        strBand = Nvl(rstmp.Fields(strBand).Value, "")
        If strShow <> "" Or strBand <> "" Then
            GetDefaultValue = strShow & "|" & strBand & "|" & i
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
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

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'���ܣ�������ʽ���߰�ť�е���һ���˵�
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.Y = vRect.Top
    vDot2.x = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.Y = vDot1.Y * 15
    vDot2.x = vDot2.x * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.Y
End Sub


Public Function CheckFormInput(objForm As Object, Optional bln������ As Boolean) As Boolean
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 And Not bln������ Then
                    MsgBox "�����д��ڷǷ��ַ���", vbInformation, App.Title
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

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
 
'------------------------------------------------------------------------------------------------
'���º������ڷ�����������ԴȨ��------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Public Function UserObject() As ADODB.Recordset
'���ܣ���ȡ��ǰ�û�������Select Ȩ�޵����б���ͼ��(�����û�������󼰱���Ȩ����)
'���أ��ɹ�=���������б�(����Ӣ˳������),ʧ��=��
'˵���������������������û�����,��ϵͳ���������ѯ
    Dim rstmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW') And USER<>'ZLSOFT'" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME and O.Owner Not in('ZLSOFT')" & _
        " Order by Sort Desc,OBJECT_NAME"
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW')" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME" & _
        " Order by Sort Desc,OBJECT_NAME"
        
    strSQL = _
        "Select Owner, Object_Name, Sign(Ascii(Object_Name) - 256) As Sort" & vbNewLine & _
        "From (Select User As Owner, Object_Name" & vbNewLine & _
        "       From User_Objects" & vbNewLine & _
        "       Where Object_Type In ('TABLE', 'VIEW')" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select Table_Schema, Table_Name" & vbNewLine & _
        "       From All_Tab_Privs" & vbNewLine & _
        "       Where Privilege = 'SELECT' And Table_Name Not Like '%_ID'" & vbNewLine & _
        "       Group By Table_Schema, Table_Name)" & vbNewLine & _
        "Order By Sort Desc, Object_Name"

    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rstmp, strSQL, "mdlPublic_UserObject")
    Set UserObject = rstmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal strArr As Variant, ByVal strCaption As String)
    'ִ�й���:
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To strArr.Count
        strSQL = strArr(i)
        Debug.Print strSQL
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    gcnOracle.CommitTrans
End Sub

