Attribute VB_Name = "mdlParseCore"
Option Explicit

'trim$Ч�ʱ�trim��
'mid$Ч�ʱ�mid��
'���ַ��Ƚ�ascw��chr��
'instr��ʹ��text��ʽЧ�ʱ�ʹ��text��ʽ��
'�ַ�����+Ч�ʵ���&
'byref���ݱ�byval�������ܸ�
'ʹ��mid��instr��ʽ����replace��ʽ�� ,ǰ���Ǳ��滻���ַ�����Ҫ���µ��ַ���������ͬ
'forѭ�����ܸ���while
'valЧ�ʵ���asc

Public Const SQL_SELECT As String = "SELECT"
Public Const SQL_FROM As String = "FROM"
Public Const SQL_WHERE As String = "WHERE"
Public Const SQL_AND As String = "AND"
Public Const SQL_OR As String = "OR"
Public Const SQL_START As String = "START"
Public Const SQL_BETWEEN As String = "BETWEEN"
Public Const SQL_ORDER As String = "ORDER"
Public Const SQL_GROUP As String = "GROUP"
Public Const SQL_HAVING As String = "HAVING"


Public Const PARSE_LEFT_BRACKET As Long = 40    '"("
Public Const PARSE_RIGHT_BRACKET As Long = 41    '")"

Public Const PARSE_LEFT_ANGLE As Long = 60      '"<"
Public Const PARSE_RIGHT_ANGLE As Long = 62     '">"

Public Const PARSE_LEFT_BRACE As Long = 123      '"{"
Public Const PARSE_RIGHT_BRACE As Long = 125     '"}"

Public Const PARSE_NULL_CHAR As Long = 32        '���ַ�

Public Enum ParallelSqlType
    pstUnion = 0
    pstMinus = 1
End Enum

'�ؼ�����Ϣ
Private Type KeyInfo
    IsStart As Boolean
    WordCount As Long
    AscTotal As Long
End Type


'******************************************************************************************************************
'******************************************************************************************************************

Public Function HasSelect(ByRef strContext As String) As Boolean
'�Ƿ����select
    Dim lngStartIndex As Long
    
    HasSelect = False
    
    lngStartIndex = InStr(strContext, SQL_SELECT)
    
    If lngStartIndex <= 0 Then Exit Function
    
    HasSelect = True
End Function


Public Function GetParNos(ByRef strContext As String, ByRef aryParNo() As Long) As Boolean
'�Ƿ��������
'aryParNo,�洢��Ӧ�Ĳ�����
    Dim lngStartIndex As Long
    Dim lngEndEndex As Long
    Dim lngUbound As Long
    Dim lngCurParNo As Long
    
    Dim strTmp As String
    
    GetParNos = False
    ReDim aryParNo(0)
    
    lngStartIndex = InStr(strContext, "[")
    
    If lngStartIndex <= 0 Then Exit Function
    
    lngEndEndex = InStr(lngStartIndex, strContext, "]")
    If lngEndEndex <= 0 Then Exit Function
    
    '������Ŀ�а����Ĳ�����
    While lngStartIndex > 0 And lngEndEndex > lngStartIndex
        strTmp = Mid$(strContext, lngStartIndex + 1, lngEndEndex - lngStartIndex - 1)
        lngCurParNo = Val(strTmp)
        
        If lngCurParNo > 0 Then
            lngUbound = UBound(aryParNo) + 1
            ReDim Preserve aryParNo(lngUbound)
            
            aryParNo(lngUbound) = lngCurParNo
        End If
        
        lngStartIndex = InStr(lngEndEndex, strContext, "[")
        If lngStartIndex > 0 Then lngEndEndex = InStr(lngStartIndex, strContext, "]")
    Wend
    
    If lngUbound <= 0 Then Exit Function
    
    GetParNos = True
End Function

Public Function RestoreBracketContext(ByVal strContext As String, _
                                        ByRef objRootBrack As clsSqlBracket, _
                                        Optional ByVal blnRestorePar As Boolean = False, _
                                        Optional ByVal blnReplaceAll As Boolean = False) As String
'�ָ������е�����
'blnRestorePar:�Ƿ�ָ�����
'blnReplaceAll:�Ƿ�ָ���������
On Error GoTo errHandle
    Dim strBrackPath() As String
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
    Dim lngUbound As Long
    Dim i As Long
    Dim j As Long
    Dim objCurBrack As clsSqlBracket
    Dim strTmp As String
    Dim strNew As String
    
    RestoreBracketContext = strContext
    
    lngStartIndex = InStr(strContext, "{%")
    lngEndIndex = InStr(lngStartIndex + 2, strContext, "}")
    
    ReDim strBrackPath(0)
    
    '��ȡ��ǰ���ݵ��滻������{%0#1},{%0#2}
    While lngStartIndex > 0 And lngEndIndex > lngStartIndex
        lngUbound = UBound(strBrackPath) + 1
        ReDim Preserve strBrackPath(lngUbound)
        
        strBrackPath(lngUbound) = Mid$(strContext, lngStartIndex, lngEndIndex - lngStartIndex + 1)
        
        lngStartIndex = InStr(lngEndIndex, strContext, "{%")
        If lngStartIndex > 0 Then
            lngEndIndex = InStr(lngStartIndex + 2, strContext, "}")
        End If
    Wend
    
    '�Ե�ǰ���ݽ����滻
    For i = 1 To UBound(strBrackPath)
        strNew = strBrackPath(i)
        
        Set objCurBrack = objRootBrack.GetBracket(strNew)
        strTmp = objCurBrack.Context
        
        '�ָ���������[1]��ʽ�ָ�Ϊ[����1]��ʽ
        If blnRestorePar Then
            For j = 1 To objCurBrack.ParReplaceCount
                strTmp = Replace(strTmp, objCurBrack.ParReplace(j), objCurBrack.ParNames(j))
            Next j
        End If
            
        If i = 1 Then
            If Len(strContext) = Len(strNew) Then
                RestoreBracketContext = Mid$(strTmp, 2, Len(strTmp) - 2)
            Else
                RestoreBracketContext = Replace(RestoreBracketContext, strNew, strTmp)
            End If
        Else
            RestoreBracketContext = Replace(RestoreBracketContext, strNew, strTmp)
        End If
        
        '�ݹ���������滻
        If blnReplaceAll Then RestoreBracketContext = RestoreBracketContext(RestoreBracketContext, objRootBrack, blnRestorePar, blnReplaceAll)
    Next i
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.RestoreBracketContext", "[RestoreBracketContext]�������>>" + vbCrLf + "  �������Ϊ��" + strContext + vbCrLf + Err.Description
    Resume
End Function


Public Sub GetBracketDetail(ByVal strContext As String, _
                            ByRef objRootBrack As clsSqlBracket, _
                            ByRef blnHasPar As Boolean, _
                            ByRef blnHasSelect As Boolean, _
                            ByRef arySql() As String, _
                            ByRef aryFunc() As String, _
                            ByRef arySqlTag() As String, _
                            ByRef aryFuncTag() As String, _
                            ByRef aryFuncLink() As Boolean)
'��ȡ�����е���ϸ��������Ƿ�����������Ƿ������ѯ��
'�� where id=2 and (���={%0#1} or ����={%0#2})
On Error GoTo errHandle
    Dim strBrackPath() As String
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
    Dim lngUbound As Long
    Dim lngBrackCount As Long
    
    Dim i As Long
    Dim objCurBrack As clsSqlBracket
    Dim strTmp As String
    Dim strNew As String
    
    blnHasPar = False
    blnHasSelect = False
    
    lngStartIndex = InStr(strContext, "{%")
    lngEndIndex = InStr(lngStartIndex + 2, strContext, "}")
    
    ReDim strBrackPath(0)
    ReDim arySql(0)
    ReDim aryFunc(0)
    ReDim arySqlTag(0)
    ReDim aryFuncTag(0)
    ReDim aryFuncLink(0)
    
    '��ȡ��ǰ���ݵ��滻������{%0#1},{%0#2}
    While lngStartIndex > 0 And lngEndIndex > lngStartIndex
        lngUbound = UBound(strBrackPath) + 1
        ReDim Preserve strBrackPath(lngUbound)
        
        strBrackPath(lngUbound) = Mid$(strContext, lngStartIndex, lngEndIndex - lngStartIndex + 1)
        
        lngStartIndex = InStr(lngEndIndex, strContext, "{%")
        If lngStartIndex > 0 Then
            lngEndIndex = InStr(lngStartIndex + 2, strContext, "}")
        End If
    Wend
    
    '�Ե�ǰ���ݽ����滻
    lngBrackCount = UBound(strBrackPath)
    For i = 1 To lngBrackCount
        strNew = strBrackPath(i)
        
        Set objCurBrack = objRootBrack.GetBracket(strNew)
        strTmp = objCurBrack.Context
        
        If objCurBrack.IsParameter Or objCurBrack.HasSubParameter Then blnHasPar = True
        
        If objCurBrack.IsSelect Then
            blnHasSelect = True
            
            lngUbound = UBound(arySql) + 1
            
            ReDim Preserve arySql(lngUbound)
            arySql(lngUbound) = Mid$(strTmp, 2, Len(strTmp) - 2)
            
            ReDim Preserve arySqlTag(lngUbound)
            arySqlTag(lngUbound) = strNew
        Else
            lngUbound = UBound(aryFunc) + 1
            
            ReDim Preserve aryFunc(lngUbound)
            aryFunc(lngUbound) = strTmp
            
            ReDim Preserve aryFuncTag(lngUbound)
            aryFuncTag(lngUbound) = strNew
            
            
            ReDim Preserve aryFuncLink(lngUbound)
            aryFuncLink(lngUbound) = False
            
            '�ж�or��and, between, start���
            If objCurBrack.OrCount > 0 Then
                aryFuncLink(lngUbound) = True
            ElseIf objCurBrack.AndCount > objCurBrack.BetweenCount Then
                aryFuncLink(lngUbound) = True
            ElseIf objCurBrack.StartCount > 0 Then
                '������or��and
                '�硰����=[����] start with id=1 connect by ... prior��
                lngEndIndex = InStr(strTmp, SQL_START)
                strTmp = Trim$(Mid$(strTmp, 1, lngEndIndex - 1))
                
                If strTmp <> "" And strTmp <> "(" And strTmp <> ")" Then
                    aryFuncLink(lngUbound) = True
                End If
            End If

        End If
    Next i
Exit Sub
errHandle:
    Err.Raise -1, "mdlParseCore.GetBracketDetail", "[GetBracketDetail]�������>>" + vbCrLf + "  �������Ϊ��" + strContext + vbCrLf + Err.Description
    Resume
End Sub

Public Function ResolveBracket(ByVal strFormatSql As String, _
    ByRef objRootBrack As clsSqlBracket, _
    ByRef strPars() As String) As String ', ByRef lngParPos() As Long
'�ֽ�����
'����������ƥ��������������������νṹ
    
On Error GoTo errHandle
    Dim strResult As String
    Dim objSubBrack As clsSqlBracket
    Dim objParentBrack As clsSqlBracket
    Dim objCurBrack As clsSqlBracket
    
    Dim strTmp As String
    Dim strMatch As String
    
    Dim strChar As String
    Dim lngLastAscw As Long
    Dim lngAscw As Long
    Dim lngCursor As Long
    
    Dim kiSelect As KeyInfo
    Dim kiOr As KeyInfo
    Dim kiAnd As KeyInfo
    Dim kiStart As KeyInfo
    Dim kiBetween As KeyInfo
    
    Dim lngParCount As Long
    Dim lngParStartIndex As Long
    Dim lngBound As Long
    Dim lngTmp As Long
    Dim blnIsParNo As Boolean
    
    Dim lngLeftTotal As Long
    Dim lngRightTotal As Long
    
    strResult = strFormatSql
    
    ReDim strPars(0)
'    ReDim lngParPos(0)
    
    Set objRootBrack = New clsSqlBracket
    Set objRootBrack.Parent = Nothing
    
    objRootBrack.Depth = 0
    objRootBrack.DepthTag = "0"
    objRootBrack.Start = 1
    
    Set objParentBrack = objRootBrack
    
    lngParCount = 0
    lngLeftTotal = 0
    lngRightTotal = 0
     
    lngLastAscw = 0
    For lngCursor = 1 To 1000000
        strChar = Mid$(strResult, lngCursor, 1)
        If Len(strChar) = 0 Then Exit For   '����ѭ������
        
        lngAscw = AscW(strChar)
        
        If lngAscw > PARSE_NULL_CHAR Then
            If lngAscw = PARSE_LEFT_BRACKET Then    '( ��������ʼ
                 
                lngLeftTotal = lngLeftTotal + 1
                
                Set objSubBrack = New clsSqlBracket
                Set objSubBrack.Parent = objParentBrack
                
                objSubBrack.Depth = objParentBrack.Depth + 1
                objSubBrack.DepthTag = objParentBrack.DepthTag & "#" & (objParentBrack.SubItemCount + 1)
                objSubBrack.Start = lngCursor
                
                Call objParentBrack.AddSubItems(objSubBrack)
                
                Set objParentBrack = objSubBrack
                
            ElseIf lngAscw = PARSE_RIGHT_BRACKET Then   ')����������
                lngRightTotal = lngRightTotal + 1
                
                strMatch = "{%" & objParentBrack.DepthTag & "}"
                '�滻����
                strTmp = Mid$(strResult, objParentBrack.Start, lngCursor - objParentBrack.Start + 1)
                strResult = Replace(strResult, strTmp, strMatch, 1, 1)
                lngTmp = Len(strMatch)
                
                lngCursor = objParentBrack.Start + lngTmp - 1   'next lngCursorʱ�����Զ���lngCursor��1
                objParentBrack.Context = strTmp
    
                
                Set objParentBrack = objParentBrack.Parent
            Else
                If lngAscw = 83 Then    'Select ���� start
                    '������select��s
                    kiSelect.IsStart = True
                    kiStart.IsStart = True
                    
                    If lngLastAscw > PARSE_NULL_CHAR Then
                        If lngLastAscw <> PARSE_LEFT_BRACKET _
                            And lngLastAscw <> PARSE_RIGHT_BRACKET Then
                            kiSelect.IsStart = False
                            kiStart.IsStart = False
                        End If
                    End If
                    
                    kiSelect.WordCount = 0
                    kiSelect.AscTotal = 0
                    
                    kiStart.WordCount = 0
                    kiStart.AscTotal = 0
                    
                ElseIf lngAscw = 79 Then 'Or
                     
                    kiOr.IsStart = True
                    If lngLastAscw > PARSE_NULL_CHAR Then
                        If lngLastAscw <> PARSE_LEFT_BRACKET _
                            And lngLastAscw <> PARSE_RIGHT_BRACKET Then
                            kiOr.IsStart = False
                        End If
                    End If
                    
                    kiOr.WordCount = 0
                    kiOr.AscTotal = 0
                    
                ElseIf lngAscw = 65 Then    'And
                
                    kiAnd.IsStart = True
                    If lngLastAscw > PARSE_NULL_CHAR Then
                        If lngLastAscw <> PARSE_LEFT_BRACKET _
                            And lngLastAscw <> PARSE_RIGHT_BRACKET Then
                            kiAnd.IsStart = False
                        End If
                    End If
                    
                    kiAnd.WordCount = 0
                    kiAnd.AscTotal = 0
                    
                ElseIf lngAscw = 66 Then    'Between
                
                    kiBetween.IsStart = True
                    If lngLastAscw > PARSE_NULL_CHAR Then
                        If lngLastAscw <> PARSE_LEFT_BRACKET _
                            And lngLastAscw <> PARSE_RIGHT_BRACKET Then
                            kiBetween.IsStart = False
                        End If
                    End If
                    
                    kiBetween.WordCount = 0
                    kiBetween.AscTotal = 0
                    
                ElseIf lngAscw = 91 Then  '"[" ������ʼ
                    lngParStartIndex = lngCursor
                    kiSelect.IsStart = False
                    
                ElseIf lngAscw = 93 Then  '"]" ��������
                    kiSelect.IsStart = False
                    
                    If lngParStartIndex > 0 Then    '���������ʽ�Ĵ���
                        strTmp = Mid$(strResult, lngParStartIndex + 1, lngCursor - lngParStartIndex - 1)
                        lngTmp = Val(strTmp)
                        
                        '�ж��Ƿ��Ѿ��Ǵ�����Ĳ�����
                        blnIsParNo = True
                        If lngTmp <= 0 Then
                            '������ʹ��val��ֵΪ0����Ҫ�ж��ַ�Ϊ0�����
                            If Len(strTmp) > 0 Then
                                If Asc(strTmp) <> 48 Then
                                    blnIsParNo = False
                                Else
                                    blnIsParNo = True
                                End If
                            Else
                                blnIsParNo = True
                            End If
                        End If
                        
                        
                        If blnIsParNo = False Then
                            lngParCount = lngParCount + 1
                            
                            strTmp = "[" & strTmp & "]"
                            '�жϸò����Ƿ��Ѿ����д���,�硰[12]����ʾ�Ѿ�������
                            lngBound = UBound(strPars) + 1
                            ReDim Preserve strPars(lngBound)
                            
                            strPars(lngBound) = strTmp
                            
'                            ReDim Preserve lngParPos(UBound(lngParPos) + 1)
'                            lngParPos(UBound(lngParPos)) = lngParCount
                            
                            strMatch = "[" & lngParCount & "]"
                            
                            strResult = Replace(strResult, strTmp, strMatch, 1)
                            lngCursor = lngParStartIndex + Len(strMatch) - 1
                            
                        Else
                            If lngTmp <= lngParCount Then
                                strTmp = strPars(lngTmp)
                                strMatch = "[" & lngTmp & "]"
                            Else
                                strTmp = "[" & strTmp & "]"
                                strMatch = strTmp
                            End If
                            
'                            ReDim Preserve lngParPos(UBound(lngParPos) + 1)
'                            lngParPos(UBound(lngParPos)) = lngTmp
                        End If
                        
                        '���ö�Ӧ�Ĳ�����
                        objParentBrack.IsParameter = True
                        Call objParentBrack.AddParLink(strTmp, strMatch)
                        
                        Set objCurBrack = objParentBrack.Parent
                        While Not objCurBrack Is Nothing
                            objCurBrack.HasSubParameter = True
                            Set objCurBrack = objCurBrack.Parent
                        Wend
                    End If
                    
                    lngParStartIndex = -1
                End If
                
                If kiSelect.IsStart Then
                    
                    kiSelect.WordCount = kiSelect.WordCount + 1
                    kiSelect.AscTotal = kiSelect.WordCount * lngAscw + kiSelect.AscTotal
                
                    If kiSelect.WordCount >= 6 Then
                        If kiSelect.AscTotal = 1564 Then    'SELECTֵ:83* 1 +  69*2 +  76*3 + 69*4 + 67*5 + 84*6

                            objParentBrack.IsSelect = True
                            
                            Set objCurBrack = objParentBrack.Parent
                            While Not objCurBrack Is Nothing
                                objCurBrack.HasSubSelect = True
                                Set objCurBrack = objCurBrack.Parent
                            Wend
                        End If

                        kiSelect.IsStart = False
                        
                        kiSelect.WordCount = 0
                        kiSelect.AscTotal = 0
                    End If
                End If
                
                If kiStart.IsStart Then
                    If IsKeyInfo(strResult, kiStart, lngCursor, lngAscw, 5, 1194) Then 'Startֵ:83* 1 +  84*2 +  65*3 + 82*4 + 84*5
                        objParentBrack.StartCount = objParentBrack.StartCount + 1
                    End If
                ElseIf kiOr.IsStart Then
                    If IsKeyInfo(strResult, kiOr, lngCursor, lngAscw, 2, 243) Then 'Orֵ:79* 1 +  82*2
                        objParentBrack.OrCount = objParentBrack.OrCount + 1
                    End If
                ElseIf kiAnd.IsStart Then
                    If IsKeyInfo(strResult, kiAnd, lngCursor, lngAscw, 3, 425) Then 'Andֵ:65* 1 +  78*2 + 68 * 3
                        objParentBrack.AndCount = objParentBrack.AndCount + 1
                    End If
                ElseIf kiBetween.IsStart Then
                    If IsKeyInfo(strResult, kiBetween, lngCursor, lngAscw, 7, 2109) Then 'Betweenֵ:66* 1 +  69*2 +  84*3 + 87*4 + 69*5 + 69*6 + 78*7
                        objParentBrack.BetweenCount = objParentBrack.BetweenCount + 1
                    End If
                End If
            End If
             
        End If
        
        lngLastAscw = lngAscw
    Next lngCursor
    
    If lngLeftTotal <> lngRightTotal Then
        If MsgBox("��⵽��ѯ�е�������ƥ�䣬�Ƿ������", vbYesNo) = vbNo Then
            Err.Raise "-1", "mdlParseCore.ResolveBracket", "��⵽��ѯ����е�����ƥ����������ѯ����Ƿ���ȷ��"
        End If
    End If
    
    ResolveBracket = strResult
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.ResolveBracket", "[ResolveBracket]�������>>" + vbCrLf + "  �������Ϊ��" + strFormatSql + vbCrLf + Err.Description
    Resume
End Function

Private Function IsKeyInfo(ByRef strSource As String, ByRef ki As KeyInfo, _
                        ByVal lngCursor As Long, ByVal lngCurAscw As Long, _
                        ByVal lngWordCount As Long, ByVal lngAscValue As Long) As Boolean
'д��ؼ�����Ϣ
On Error GoTo errHandle
    Dim lngEndAsc As Long
    Dim blnIsKey As Boolean
    
    IsKeyInfo = False
    ki.WordCount = ki.WordCount + 1
    ki.AscTotal = ki.WordCount * lngCurAscw + ki.AscTotal

    If ki.WordCount >= lngWordCount Then
        If ki.AscTotal = lngAscValue Then
            blnIsKey = True
            
            lngEndAsc = AscW(Mid$(strSource, lngCursor + 1, 1))
            
            '�жϹؼ��ֺ����Ƿ�ո������
            If lngEndAsc > PARSE_NULL_CHAR Then
                If lngEndAsc <> PARSE_LEFT_BRACKET _
                    And lngEndAsc <> PARSE_RIGHT_BRACKET Then
                    blnIsKey = False
                End If
            End If
            
            If blnIsKey Then IsKeyInfo = True
        End If

        ki.IsStart = False
        
        ki.WordCount = 0
        ki.AscTotal = 0
    End If
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.IsKeyInfo", "[IsKeyInfo]�������>>" + vbCrLf + "  �������Ϊ��" + strSource + vbCrLf + Err.Description
    Resume
End Function

Public Function GetWithContext(ByVal strSource As String, ByRef strWithContext As String) As String
'������ʽ���£�
'1: with a as {%0#1}, b as {%0#2} {%0#3}
'2: with a as {%0#1}, b as {%0#2> select * from dual
'3: with a as {%0#1}, b as {%0#2> {%0#3} union all select * from dual
'4: with a as {%0#1}, b as {%0#2> {%0#3} minus select f1, f2, f3 from dual
'5: with a as {%0#1} select f1, f2, f3 from dual
'6: with a as {%0#1}, b as {%0#2} select f1, f2, f3 from dual

'����������
'��ȡ as{��ʽ��ȡas{֮���"}"λ��
'ȷ������λ�ú�ȡwith����Ч����λ��֮�������
On Error GoTo errHandle
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
    
    Dim i As Long
    Dim lngLen As Long
    Dim strChr As String
    Dim lngAscw As Long
    Dim lngTmp As Long
    
    GetWithContext = strSource
    strWithContext = ""
    
    If Len(strSource) <= 0 Then Exit Function
    
    lngStartIndex = InStr(strSource, "WITH")
    If lngStartIndex <= 0 Then Exit Function    'û��with���˳�
    
    lngLen = Len(strSource)
    lngEndIndex = InStr(lngStartIndex, strSource, "}")
    lngTmp = lngEndIndex
    
    While lngTmp > 0
        lngTmp = InStr(lngTmp + 1, strSource, "AS")
        If lngTmp > 0 Then
            For i = lngTmp + 2 To lngLen
                strChr = Mid$(strSource, i, 1)
                lngAscw = AscW(strChr)
                
                If lngAscw > PARSE_NULL_CHAR Then
                    If lngAscw = PARSE_LEFT_BRACE Then
                        lngTmp = InStr(i, strSource, "}")
                    Else
                        lngTmp = 0
                    End If
                    
                    Exit For
                End If
            Next i
        End If
        
        If lngTmp > 0 Then lngEndIndex = lngTmp
    Wend
    
    If lngStartIndex <= 0 Or lngEndIndex <= lngStartIndex Then Exit Function

    strWithContext = Mid$(strSource, lngStartIndex, lngEndIndex - lngStartIndex + 1)
    GetWithContext = Mid$(strSource, 1, lngStartIndex - 1) + Mid$(strSource, lngEndIndex + 1, lngLen)
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.GetWithContext", "[GetWithContext]�������>>" + vbCrLf + "  �������Ϊ��" + strSource + vbCrLf + Err.Description
    Resume
End Function

Public Function GetWithPart(ByVal strWithContext As String) As String()
On Error GoTo errHandle
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
     
    Dim lngUbound As Long
    Dim strResult() As String
    
    ReDim strResult(0)
    
    lngStartIndex = InStr(strWithContext, "{")
    lngEndIndex = InStr(lngStartIndex, strWithContext, "}")
    
    While lngStartIndex > 0 And lngEndIndex > lngStartIndex
        lngUbound = UBound(strResult) + 1
        ReDim Preserve strResult(lngUbound)
        
        strResult(lngUbound) = Mid$(strWithContext, lngStartIndex, lngEndIndex - lngStartIndex + 1)
        
        lngStartIndex = InStr(lngEndIndex, strWithContext, "{")
        If lngStartIndex > 0 Then
            lngEndIndex = InStr(lngStartIndex, strWithContext, "}")
        End If
    Wend
 
    GetWithPart = strResult
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.GetWithPart", "[GetWithPart]�������>>" + vbCrLf + "  �������Ϊ��" + strWithContext + vbCrLf + Err.Description
    Resume
End Function


Public Function GetParallelPart(ByVal strSource As String, ByVal lngParallelType As ParallelSqlType) As String()
'��ȡ���в��ֵ���䣬��union, minus֮��
On Error GoTo errHandle
    Dim strResult() As String
    Dim lngTmpIndex As Long
    Dim strSplitResult() As String
    
    Dim i As Long
    Dim lngUbound1 As Long
    Dim lngUbound2 As Long
    Dim strTmp As String
    
    ReDim GetParallelPart(0)
    '�ж��Ƿ���ڽ����Ĺؼ���,�ؼ���Ϊunion��minus
    strSource = Replace(strSource, "UNION ALL", "UNION")
    
    If lngParallelType = pstUnion Then
        lngTmpIndex = InStr(1, strSource, "UNION")
    Else
        lngTmpIndex = InStr(1, strSource, "MINUS")
    End If
    
    '���������union����minus���˳�
    If lngTmpIndex <= 0 Then Exit Function
    
    '�ֽ����
    If lngParallelType = pstUnion Then
        strSplitResult = Split(strSource, "UNION")
    Else
        strSplitResult = Split(strSource, "MINUS")
    End If
    
    ReDim strResult(0)
    lngUbound1 = UBound(strSplitResult)
    
    For i = 0 To lngUbound1
        strTmp = strSplitResult(i)
        
        '�ж��������Ƿ����������minus����union
        If lngParallelType = pstUnion Then
            lngTmpIndex = InStr(1, strTmp, "MINUS")
        Else
            lngTmpIndex = InStr(1, strTmp, "UNION")
        End If
        
        '������minus����union,�����0�����ڣ��ڲ���Ϊ��Ӧ���Ӿ䣬
        '�� aa minus bb union cc union dd, ������union����ʱ����0������minus�����unionֻ��cc��dd��union�����
        If lngTmpIndex <= 0 Then
            lngUbound2 = UBound(strResult) + 1

            ReDim Preserve strResult(lngUbound2)
            strResult(lngUbound2) = strTmp
        Else
            If i > 0 Then
                strTmp = Mid$(strTmp, 1, lngTmpIndex - 1)
                
                lngUbound2 = UBound(strResult) + 1
                
                ReDim Preserve strResult(lngUbound2)
                strResult(lngUbound2) = strTmp
            End If
        End If
    Next i
    
    GetParallelPart = strResult
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.GetParallelPart", "[GetParallelPart]�������>>" + vbCrLf + "  �������Ϊ��" + strSource + vbCrLf + Err.Description
    Resume
End Function


Public Function GetSelectFromPart(ByVal strSelectFromSection As String) As String()
'����select �� from�д��ڵ��Ӳ�ѯ��������
On Error GoTo errHandle
    Dim strResult() As String
    Dim strTmp As String
    Dim strCurFrom As String
    Dim lngStartIndex As Long
    Dim lngEndIndex As Long
 
    Dim lngUbound As Long


    ReDim GetSelectFromPart(0)

    If Trim$(strSelectFromSection) = "" Then
        Exit Function
    End If
 
    strCurFrom = "," & strSelectFromSection & ","
    lngStartIndex = 1
    lngEndIndex = InStr(2, strCurFrom, ",")
    
    ReDim strResult(0)
    
    While lngStartIndex > 0 And lngEndIndex > lngStartIndex
        'ȡ����֮�������
        strTmp = Mid$(strCurFrom, lngStartIndex, lngEndIndex - lngStartIndex + 1)
'        'ֻ���ذ�������������
'        If InStr(strTmp, "{%") > 0 Then
            lngUbound = UBound(strResult) + 1
            
            ReDim Preserve strResult(lngUbound)
            strResult(lngUbound) = strTmp
'        End If
        
        lngStartIndex = lngEndIndex
        lngEndIndex = InStr(lngStartIndex + 1, strCurFrom, ",")
    Wend

    GetSelectFromPart = strResult
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.GetSelectFromPart", "[GetSelectFromPart]�������:" + vbCrLf + " �������Ϊ��" + strSelectFromSection + vbCrLf + Err.Description
    Resume
End Function


Public Function GetWherePart(ByVal strFormatWhere As String) As String()
'��ȡWhere������Ŀ
'������������
'����ɨ�� And|or|start  ��λ��
'�ж�And |Or֮�������Ƿ�Ϊ�������������������ȡ��
'where���������ʽ����
'x=x and {%0#1} or <%0#2}
'{%0#1} or {%0#2} and x=x
'{%0#1}
'x=x and y={%0#1} and decode{%0#2} > 0 and z=z
'xx and y between {%0#1} and {%0#2} and z=z

On Error GoTo errHandle
    Dim lngAndIndex As Long
    Dim lngOrIndex As Long
    Dim lngStartLinkIndex As Long
    Dim lngBetweenIndex As Long
    Dim lngStartIndex As Long
    
    Dim blnValidLink As Boolean
    
    Dim strResult() As String
    Dim strTemp As String
    Dim lngSplitIndex As Long
    Dim lngSplitLen As Long
    Dim lngUbound As Long
    Dim blnMoveBetween As Boolean
    
    ReDim GetWherePart(0)
    
    If Trim$(strFormatWhere) = "" Then
        Exit Function
    End If

    
    lngStartIndex = 1
    
    lngAndIndex = InStr(strFormatWhere, SQL_AND)
    lngOrIndex = InStr(strFormatWhere, SQL_OR)
    lngStartLinkIndex = InStr(strFormatWhere, SQL_START)
    
    If lngAndIndex = 0 And lngOrIndex = 0 And lngStartLinkIndex = 0 Then
        lngAndIndex = Len(strFormatWhere) + 1
        lngStartLinkIndex = lngAndIndex
        lngOrIndex = lngAndIndex
    End If
     
    ReDim strResult(0)
    lngSplitIndex = GetMinIndex(lngAndIndex, lngOrIndex, lngStartLinkIndex, lngSplitLen)
    
    blnMoveBetween = False
    While lngStartIndex > 0 And lngSplitIndex > lngStartIndex
 
        '����and����orλ��
        '��Ҫ����������ʽ��
        '�磺" And ", " Or ", ")And "," Or(", ")AND(", ")OR(", " Or{", "}OR", "AND{" or "}AND"
        blnValidLink = IsValidLink(strFormatWhere, lngSplitIndex, lngSplitLen)
        
        If blnValidLink Then
            strTemp = Mid$(strFormatWhere, lngStartIndex, lngSplitIndex - lngStartIndex)
            
            '�ж��Ƿ����between����
            lngBetweenIndex = InStr(strTemp, SQL_BETWEEN)
            If (blnMoveBetween = False And lngBetweenIndex <> 0 And (lngStartIndex + lngBetweenIndex) < lngSplitIndex) Then
                lngSplitIndex = lngSplitIndex + Abs(lngSplitLen)
                
                blnMoveBetween = True
            Else
'                If InStr(strTemp, "{%") > 0 Then
                    lngUbound = UBound(strResult) + 1
                    
                    ReDim Preserve strResult(lngUbound)
                    strResult(lngUbound) = strTemp
'                End If
                
                '����ж�
                blnMoveBetween = False
                
                If lngSplitLen <= 0 Then
                    lngStartIndex = lngSplitIndex
                    lngSplitIndex = lngSplitIndex + Abs(lngSplitLen)
                Else
                    lngStartIndex = lngSplitIndex + Abs(lngSplitLen)
                    lngSplitIndex = lngStartIndex
                End If

            End If
        Else
            lngSplitIndex = lngSplitIndex + Abs(lngSplitLen)
        End If
        
        '����ƶ�
        lngAndIndex = InStr(lngSplitIndex, strFormatWhere, SQL_AND)
        lngOrIndex = InStr(lngSplitIndex, strFormatWhere, SQL_OR)
        lngStartLinkIndex = InStr(lngSplitIndex, strFormatWhere, SQL_START)
        
        lngSplitIndex = GetMinIndex(lngAndIndex, lngOrIndex, lngStartLinkIndex, lngSplitLen)
        
        If lngSplitIndex <= 0 Then lngSplitIndex = Len(strFormatWhere) + 1
    Wend
    
    GetWherePart = strResult
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.GetWherePart", "[GetWherePart]�������>>" + vbCrLf + "  �������Ϊ��" + strFormatWhere + vbCrLf + Err.Description
    Resume
End Function


Public Sub Parse(ByVal strFormatSourceSql As String, _
    ByRef strWithSection As String, _
    ByRef strSelectSection As String, _
    ByRef strFormSection As String, _
    ByRef strWhereSection As String, _
    ByRef strOtherSection As String, _
    ByRef strFuncSection As String)
'���������Ĳ�ѯ��䣬��������������sql��ѯ�Ѿ��޸�Ϊ�򵥵�select ... from ... where ... order ... ��ʽ
'select ... from �������
'���÷ֶ��жϷ�ʽ���ж�select ... from֮���Ƿ񻹴���select�ؼ��֣�
'������ڣ���ӵ�ǰ�����Ƶ�from֮��Ķ��䣬�ڼ����ж�from֮���Ƿ���select,���û�����ʾ�͵�һ��selectƥ��
'strWithSection ����with���ֵ����
'strSelectSection ����select���ֵ����
'strFormSection ����form���ֵ����
'strWhereSection ����where���ֵ����
'strOtherSection ����other���ֵ����
'strFuncSection ���غ������ܲ��֣�����SELECT��ѯ��
'�������ʱ����Ҫ����where�ؼ�����start with֮ǰ

On Error GoTo errHandle
    Dim lngSelectIndex As Long
    Dim lngFromIndex As Long
    Dim lngWhereIndex As Long
    Dim lngOtherIndex As Long
    Dim lngOrderIndex As Long
    Dim lngGroupIndex As Long
    Dim lngHavingIndex As Long
    Dim lngUbound As Long
    
    Dim strFormatSql As String
    Dim strTemp As String
    
    Dim lngSqlLen As Long
    Dim i As Long

    Dim blnHasWhere As Boolean
    
    
    If Len(strFormatSourceSql) <= 0 Then
        Exit Sub
    End If
    
    strFormatSql = GetWithContext(strFormatSourceSql, strWithSection)
    
    lngSqlLen = Len(strFormatSql)
    
    '����Select��λ��select ... from
    lngSelectIndex = InStr(1, strFormatSql, SQL_SELECT)
    lngFromIndex = InStr(1, strFormatSql, SQL_FROM)
    
    If lngSelectIndex <= 0 And lngFromIndex <= 0 Then
        strFuncSection = strFormatSourceSql
        strWithSection = ""
'        Err.Raise -1, "mdlParseCore.ParseSql", "������Ч��SQL��ѯ��"
        Exit Sub
    End If
    
    '��ȡselect ... from ֮�������
    strSelectSection = Mid$(strFormatSql, lngSelectIndex + 6, lngFromIndex - lngSelectIndex - 6)
    
    lngWhereIndex = InStr(lngFromIndex + 4, strFormatSql, SQL_WHERE)
    blnHasWhere = True
    If lngWhereIndex <= 0 Then
        blnHasWhere = False
        lngWhereIndex = InStr(lngFromIndex + 4, strFormatSql, SQL_START)
    End If
    
    
    '����Where���֮���Order|Group|Having��λ��,where......order
    lngOrderIndex = InStr(1, strFormatSql, SQL_ORDER)
    lngGroupIndex = InStr(1, strFormatSql, SQL_GROUP)
    lngHavingIndex = InStr(1, strFormatSql, SQL_HAVING)
    
    lngOtherIndex = GetMinValue(lngOrderIndex, lngGroupIndex, lngHavingIndex)
        
    If lngWhereIndex > 0 Then
        '��ȡfrom ... where ֮�������
        strFormSection = Mid$(strFormatSql, lngFromIndex + 4, lngWhereIndex - lngFromIndex - 4)
        
        If lngOtherIndex > 0 Then
            '��ȡwhere ... order ����
            If blnHasWhere Then
                strWhereSection = Mid$(strFormatSql, lngWhereIndex + 5, lngOtherIndex - lngWhereIndex - 5)
            Else
                strWhereSection = Mid$(strFormatSql, lngWhereIndex, lngOtherIndex - lngWhereIndex)
            End If
        Else
            If blnHasWhere Then
                strWhereSection = Mid$(strFormatSql, lngWhereIndex + 5, lngSqlLen + 1 - lngWhereIndex - 5)
            Else
                strWhereSection = Mid$(strFormatSql, lngWhereIndex, lngSqlLen + 1 - lngWhereIndex)
            End If
        End If
    Else
        strWhereSection = ""
        
        '��ȡfrom ... order ֮�������
        If lngOtherIndex > 0 Then
            strFormSection = Mid$(strFormatSql, lngFromIndex + 4, lngOtherIndex - lngFromIndex - 4)
        Else
            strFormSection = Mid$(strFormatSql, lngFromIndex + 4, lngSqlLen + 1 - lngFromIndex - 4)
        End If
    End If
    
    '��ȡother����
    If lngOtherIndex = 0 Or lngOtherIndex >= lngSqlLen Then
        strOtherSection = ""
    Else
        strOtherSection = Mid$(strFormatSql, lngOtherIndex, lngSqlLen + 1 - lngOtherIndex)
    End If
    
Exit Sub
errHandle:
    Err.Raise -1, "mdlParseCore.Parse", "[Parse]�������>>" + vbCrLf + "  �������Ϊ��" + strFormatSql + vbCrLf + Err.Description
    Resume
End Sub

Public Function FormatSql(ByVal strSql As String, ByRef strQuotes() As String) As String
'��ʽ��sql���
On Error GoTo errHandle
    Dim strSourceSql As String
    
    strSourceSql = strSql
    
    strSourceSql = FormatEnter(strSql)
    strSourceSql = FormatSymbolsInQuote(strSourceSql, strQuotes)
    strSourceSql = FormatNullChar(strSourceSql)
    strSourceSql = UCase(strSourceSql)
    
    '����ո��滻Ϊ���ո�
'    strSourceSql = FormatBracket(strSourceSql)
    
    FormatSql = strSourceSql
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.FormatSql", "[FormatSql]�������>>" + vbCrLf + "  �������Ϊ��" + strSql + vbCrLf + Err.Description
    Resume
End Function


Public Function RestoreQuote(ByVal strSql As String, ByRef strQuotes() As String) As String
'�ָ�˫����֮�������
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = UBound(strQuotes)
    
    RestoreQuote = strSql
    For i = 1 To lngCount
        RestoreQuote = Replace(RestoreQuote, "<@" & i & "/>", strQuotes(i))
    Next i
End Function

Public Function RestoreParIndex(ByVal strSql As String, ByRef strPars() As String) As String
'�ָ���������
    Dim i As Long
    Dim lngCount As Long
    
    RestoreParIndex = strSql
    lngCount = UBound(strPars)
    
    For i = 1 To lngCount
        RestoreParIndex = Replace(RestoreParIndex, strPars(i), "[" & i & "]")
    Next i
End Function


'====================================================================================================================
'********************************************************************************************************************
'====================================================================================================================

Private Function FormatSymbolsInQuote(ByVal strSource As String, ByRef strQuotes() As String) As String
'��ʽ���������ڵ�����
On Error GoTo errHandle
    Dim i As Long
    Dim regxpStr As Object
    Dim objMatchs As Object
    Dim blnAllowReplace As Boolean
    Dim strReplaces As String
    Dim lngUbound As Long
    
    FormatSymbolsInQuote = strSource
    ReDim strQuotes(0)
    
    If InStr(FormatSymbolsInQuote, "'") <= 0 Then Exit Function
    
     
    Set regxpStr = CreateObject("VBScript.RegExp")
    
    regxpStr.Global = True
    regxpStr.IgnoreCase = True
    regxpStr.MultiLine = True
    
    '�ж��Ƿ���ڵ����ţ�������ڣ��򽫵����ŵ����ݽ����滻,��Ϊ�������ڿ��ܳ��֡����������������ַ�,��Nvl(instr('''a123)'',( 456)xx x''b123'',', a.�걾��λ),1)>a.�걾��λ  ')',
    '���������е�˫�����滻Ϊ�����ַ�
    regxpStr.Pattern = "\''"    '"\''(?=[^\'])"    '\''[^\']
    FormatSymbolsInQuote = regxpStr.Replace(FormatSymbolsInQuote, "<@Quote2>")    '�滻�����£�Nvl(instr('||a123)||,( 456)xx x||b123||,', a.�걾��λ),1)>a.�걾��λ  ')',
    
    '�ٽ��������ڵ������滻Ϊ�����ַ�
    regxpStr.Pattern = "\'(.*?)\'"
    Set objMatchs = regxpStr.Execute(FormatSymbolsInQuote)
    
    strReplaces = ""
    If objMatchs.Count > 0 Then
        For i = 0 To objMatchs.Count - 1
            blnAllowReplace = True
'            If objMatchs(i).Length <= 3 Then
'                If InStr("(),", Replace(objMatchs(i), "'", "")) >= 1 Then
'                    blnAllowReplace = True
'                Else
'                    blnAllowReplace = False
'                End If
'            End If

'            If IsNumeric(objMatchs(i)) Then blnAllowReplace = False
            
            If blnAllowReplace Then
                If InStr(strReplaces, objMatchs(i)) <= 0 Then
                    lngUbound = UBound(strQuotes) + 1
                    
                    ReDim Preserve strQuotes(lngUbound)
                    strQuotes(lngUbound) = Replace(objMatchs(i), "<@Quote2>", "''")
                    
                    FormatSymbolsInQuote = Replace(FormatSymbolsInQuote, objMatchs(i), "<@" & lngUbound & "/>")
                    
                    strReplaces = strReplaces & "," & objMatchs(i) & ","
                End If
            End If
        Next i
    End If
    
    FormatSymbolsInQuote = Replace(FormatSymbolsInQuote, "<@Quote2>", "''")
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.FormatSymbolsInQuote", "[FormatSymbolsInQuote]�������>>" + vbCrLf + "  �������Ϊ��" + strSource + vbCrLf + Err.Description
End Function

Private Function FormatNullChar(ByVal strSql As String) As String
'��ʽ�����ַ�,������գ���ʽ��Ϊһ����
On Error GoTo errHandle
    FormatNullChar = Replace(strSql, vbTab, " ")
    
    '�滻����4���������ո�
    While InStr(FormatNullChar, "    ") > 0
        FormatNullChar = Replace(FormatNullChar, "    ", " ")
    Wend
    
    '�滻����2���������ո�
    While InStr(FormatNullChar, "  ") > 0
        FormatNullChar = Replace(FormatNullChar, "  ", " ")
    Wend
    
    FormatNullChar = Replace(FormatNullChar, "( ", "(")
    FormatNullChar = Replace(FormatNullChar, " )", ")")
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.FormatNullChar", "[FormatNullChar]�������>>" + vbCrLf + "  �������Ϊ��" + strSql + vbCrLf + Err.Description
    Resume
End Function

Private Function FormatEnter(ByVal strSource As String) As String
'��ʽ��sql����еĻس�����
    FormatEnter = Replace(Replace(strSource, vbCr, " "), vbLf, " ")
End Function


Private Function IsValidLink(ByRef strWhere As String, ByVal lngLinkIndex As Long, ByVal lngAdjustLen As Long) As Boolean
'�ж��Ƿ���Ч���������ӷ��� �������ӷ�ָ or, and ,start
On Error GoTo errHandle
    Dim strChr As String
    Dim lngLinkLen As Long
    Dim lngAscw As Long
    
    IsValidLink = True
    If lngLinkIndex >= Len(strWhere) Then Exit Function
    
    lngLinkLen = lngAdjustLen
    If lngLinkLen <= 0 Then lngLinkLen = 5  'start �ĳ���
    
    If lngLinkIndex > 0 Then
        If lngLinkIndex > 1 Then
            strChr = Mid$(strWhere, lngLinkIndex - 1, 1)
            lngAscw = AscW(strChr)
            If lngAscw > PARSE_NULL_CHAR And lngAscw <> PARSE_RIGHT_BRACKET And lngAscw <> PARSE_RIGHT_BRACE Then   'Ascll��41��ʾ")"
                '������and,or��ʽ
                IsValidLink = False
                Exit Function
            End If
        End If
        
        strChr = Mid$(strWhere, lngLinkIndex + lngLinkLen, 1)
        lngAscw = AscW(strChr)
        If lngAscw > PARSE_NULL_CHAR And lngAscw <> PARSE_LEFT_BRACKET And lngAscw <> PARSE_LEFT_BRACE Then  'Ascll��40��ʾ"("
            '������and,or��ʽ
            IsValidLink = False
            Exit Function
        End If
    End If
Exit Function
errHandle:
    Err.Raise -1, "mdlParseCore.IsValidLink", "[IsValidLink]�������>>" + vbCrLf + "  �������Ϊ��" + strWhere + vbCrLf + Err.Description
    Resume
End Function


Private Function GetMinIndex(ByVal lngAndIndex As Long, ByVal lngOrIndex As Long, _
                        ByVal lngStartIndex As Long, ByRef lngSplitLen As Long) As Long
'��ȡ��ǰ���and��or��start���ڵ�����
    Dim lngMin As Long
    Dim lngV1 As Long
    Dim lngV2 As Long
    Dim lngV3 As Long
    
    lngMin = 100000

    lngV1 = IIf(lngAndIndex > 0, lngAndIndex, 100000)
    lngV2 = IIf(lngOrIndex > 0, lngOrIndex, 100000)
    lngV3 = IIf(lngStartIndex > 0, lngStartIndex, 100000)
    
    If lngMin > lngV1 Then  'and
        lngMin = lngV1
        lngSplitLen = 3
    End If
    
    If lngMin > lngV2 Then  'or
        lngMin = lngV2
        lngSplitLen = 2
    End If
    
    If lngMin > lngV3 Then  'start
        lngMin = lngV3
        lngSplitLen = -5
    End If

    If lngMin = 100000 Then lngMin = 0

    GetMinIndex = lngMin
End Function

