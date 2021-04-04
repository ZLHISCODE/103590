Attribute VB_Name = "mdlDiff"
Option Explicit
'ģ��˵��:�����ı�����Ա�

Private Enum ��ɫ
    ��ɫ = &HFFFFFF
    ����ɫ = &HC9C9CD
    ��ɫ = &H106E2A
    ��ɫ = &H0&
    ��ɫ = &H4040FF
    ��ɫ = vbBlue
End Enum

Public Sub MergeDiff(strTxt1 As String, strTxt2 As String)
    '����CompareIt�����Աȹ�����ַ��п���,���д���ϲ�
    Dim i As Long, j As Long, blnDelete As Boolean
    Dim lngS As Long, lngE As Long
    Dim arrTxt1() As String, arrTxt2() As String
    
    arrTxt1 = Split(strTxt1, vbNewLine)
    arrTxt2 = Split(strTxt2, vbNewLine)
    
    Do While i < UBound(arrTxt1)
        '������촦
        '˵��:�ڵ���CompaIt������,���ڲ�ͬ���ı�,���ж���Ŀ���,�����������Ŀ���
        If arrTxt1(i) = "" And arrTxt2(i) <> "" Then
            If lngS = 0 Then lngS = i
            lngE = lngE + 1
        Else
            For j = lngS To lngS + lngE - 1
                If arrTxt1(j) = "" And arrTxt2(j) <> "" And arrTxt1(j + lngE) <> "" And arrTxt2(j + lngE) = "" Then
                    blnDelete = True
                Else
                    blnDelete = False
                    Exit For
                End If
            Next
            
            If blnDelete Then
                For j = 1 To lngE
                    DelElmentByIdx arrTxt1, lngS
                    DelElmentByIdx arrTxt2, lngS + lngE
                    i = i - 1 'ɾ����Ԫ��,�±�-1
                Next
            End If
            lngS = 0: lngE = 0
        End If
        i = i + 1
    Loop
    
    strTxt1 = "": strTxt2 = ""
    
    strTxt1 = GetStrFromArr(arrTxt1, 0, UBound(arrTxt1))
    strTxt2 = GetStrFromArr(arrTxt2, 0, UBound(arrTxt2))
End Sub

Public Sub CompareIt(ByRef strTxt1 As String, ByRef strTxt2 As String)
    '����:�ԱȲ���󽫸�ʽ�����ı�����
    Dim arrTxt1() As String, arrTxt2() As String
    Dim strPrefix As String, strSuffix As String
    Dim strBefore1 As String, strAfter1 As String
    Dim strBefore2 As String, strAfter2 As String
    Dim strMiddle As String
    Dim i As Long, j As Long, lngRow As Long
    
    If UCase(TrimEx(strTxt1)) = UCase(TrimEx(strTxt2)) Then
        Exit Sub '��ͬ���ı�ֱ�ӷ���
    End If
    
    arrTxt1 = Split(strTxt1, vbNewLine)
    arrTxt2 = Split(strTxt2, vbNewLine)
    
    '��ȡͷ������ͬ����
    strPrefix = GetCommonPrefix(arrTxt1, arrTxt2, i)
    
    '��ȡβ������ͬ����
    strSuffix = GetCommonSuffix(arrTxt1, arrTxt2, j)
    
    '��ȥͷβ����ͬ����
    If strPrefix <> "" Or strSuffix <> "" Then
        strTxt1 = GetStrFromArr(arrTxt1, i, UBound(arrTxt1) - j)
        strTxt2 = GetStrFromArr(arrTxt2, i, UBound(arrTxt2) - j)
    End If
    
    arrTxt1 = Split(strTxt1, vbNewLine)
    arrTxt2 = Split(strTxt2, vbNewLine)
    If Replace(strTxt1, vbNewLine, "") = "" Or Replace(strTxt2, vbNewLine, "") = "" Then
        If Replace(strTxt1, vbNewLine, "") = "" Then
            For i = 1 To UBound(arrTxt2)
                strTxt1 = IIf(i = 1, vbNewLine, strTxt1 & vbNewLine)
            Next
        ElseIf Replace(strTxt2, vbNewLine, "") = "" Then
            For i = 1 To UBound(arrTxt1)
                strTxt2 = IIf(i = 1, vbNewLine, strTxt2 & vbNewLine)
            Next
        End If
        
        strTxt1 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & strTxt1 & IIf(strSuffix = "", "", vbNewLine & strSuffix)
        strTxt2 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & strTxt2 & IIf(strSuffix = "", "", vbNewLine & strSuffix)
        Exit Sub
    End If
    
    '��ȡ��ͬ���ж�
    strMiddle = GetCommonMid(arrTxt1, arrTxt2)
    
    If Replace(strMiddle, " ", "") = "" Then
        CheckDiff strTxt1, strTxt2
        
        strTxt1 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & strTxt1 & IIf(strSuffix = "", "", vbNewLine & strSuffix)
        strTxt2 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & strTxt2 & IIf(strSuffix = "", "", vbNewLine & strSuffix)
    Else
        '����ͬ���м��,ȥ���м䲿��,��ǰ�����ν��е���
        GetBeforeAndAfter arrTxt1, Split(strMiddle, vbNewLine), strBefore1, strAfter1
        GetBeforeAndAfter arrTxt2, Split(strMiddle, vbNewLine), strBefore2, strAfter2
        
        strTxt1 = strMiddle
        strTxt2 = strMiddle
        CompareIt strBefore1, strBefore2
        CompareIt strAfter1, strAfter2
        
        strTxt1 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & IIf(strBefore1 = "", "", strBefore1 & vbNewLine) _
                        & strTxt1 & IIf(strAfter1 = "", "", vbNewLine & strAfter1) & IIf(strSuffix = "", "", vbNewLine & strSuffix)
                        
        strTxt2 = IIf(strPrefix = "", "", strPrefix & vbNewLine) & IIf(strBefore2 = "", "", strBefore2 & vbNewLine) _
                        & strTxt2 & IIf(strAfter2 = "", "", vbNewLine & strAfter2) & IIf(strSuffix = "", "", vbNewLine & strSuffix)
        
    End If
    
End Sub



Private Function GetCommonPrefix(arrTxt1 As Variant, arrTxt2 As Variant, Optional ByRef lngIdx As Long) As String
    '��ȡ������ǰ׺���ı�
    Dim lngMin As Long, i As Long, j As Long
    Dim strPrefix As String
    
    i = UBound(arrTxt1): j = UBound(arrTxt2)
    lngMin = IIf(i > j, j, i)
    
    For i = 0 To lngMin
       If UCase(TrimEx(arrTxt1(i))) <> UCase(TrimEx(arrTxt2(i))) Then
            Exit For
        Else
            If i = 0 Then
                strPrefix = arrTxt1(i)
            Else
                strPrefix = strPrefix & vbNewLine & arrTxt1(i)
            End If
        End If
    Next
    
    lngIdx = i
    GetCommonPrefix = strPrefix
End Function


Private Function GetCommonSuffix(ByVal arrTxt1 As Variant, ByVal arrTxt2 As Variant, Optional ByRef lngIdx As Long) As String
    '��ȡ������׺���ı�
    Dim lngMin As Long, i As Long, j As Long
    Dim strSuffix As String
    
    i = UBound(arrTxt1): j = UBound(arrTxt2)
    lngMin = IIf(i < j, i, j)
    
    For i = 0 To lngMin
        If TrimEx(UCase(arrTxt1(UBound(arrTxt1) - i))) <> TrimEx(UCase(arrTxt2(UBound(arrTxt2) - i))) Then
            Exit For
        Else
            If i = 0 Then
                strSuffix = arrTxt1(UBound(arrTxt1) - i)
            Else
                strSuffix = arrTxt1(UBound(arrTxt1) - i) & vbNewLine & strSuffix
            End If
        End If
    Next
    
    lngIdx = i
    GetCommonSuffix = strSuffix
End Function


Private Function GetStrFromArr(arrTxt As Variant, ByVal lngS As Long, ByVal lngE As Long, Optional ByVal blnTrim As Boolean) As String
    '����:���ظ��ݴ�������鿪ʼ�±�ͽ����±�,�����ַ���
    Dim i As Long, strResult As String

    If lngS < 0 Then Exit Function
    If lngE > UBound(arrTxt) Then lngE = UBound(arrTxt)
    
    For i = lngS To lngE
        If i = lngS Then
            strResult = IIf(blnTrim, UCase(TrimEx(arrTxt(i))), arrTxt(i))
        Else
            strResult = strResult & vbNewLine & IIf(blnTrim, UCase(TrimEx(arrTxt(i))), arrTxt(i))
        End If
    Next
    GetStrFromArr = strResult
End Function


Private Function GetCommonMid(arrTxt1 As Variant, arrTxt2 As Variant) As String
    '��ȡ�����м��
    Dim strTxt1 As String, strTxt2 As String
    Dim arrBig() As String, arrSmall() As String
    Dim strBig As String, strSmall As String
    Dim i As Long, j As Long, strCommon As String
    Dim lngStep As Long, lngLen As Long
    
    i = UBound(arrTxt1): j = UBound(arrTxt2)
    If i < 50 Or j < 50 Then '����ı�����������(С��50��),�Ͳ���Ҫ��ȡ�����м��
        Exit Function
    End If

    If i < j Then
        arrSmall = arrTxt1
        strBig = GetStrFromArr(arrTxt2, 1, j - 1, True)
    Else
        arrSmall = arrTxt2
        strBig = GetStrFromArr(arrTxt1, 1, i - 1, True)
    End If

    '�ƶ���С�ַ�����1/16�Ӵ�����ƥ��
    lngStep = (UBound(arrSmall) - 1) \ 32
    lngLen = (UBound(arrSmall) - 1) \ 16
    
    For i = 0 To UBound(arrSmall) - 1 Step lngStep
        strTxt1 = GetStrFromArr(arrSmall, i + 1, i + lngLen - 1, True)
        strTxt2 = GetStrFromArr(arrSmall, i + 1, i + lngLen - 1)
        
        'ȥ��ǰ�����
        Do While Right(strTxt2, 1) = vbCr Or Right(strTxt2, 1) = vbLf
            strTxt1 = Left(strTxt1, Len(strTxt1) - 1)
            strTxt2 = Left(strTxt2, Len(strTxt2) - 1)
        Loop
        
        Do While Left(strTxt2, 1) = vbLf Or Left(strTxt2, 1) = vbCr
            strTxt1 = Right(strTxt1, Len(strTxt1) - 1)
            strTxt2 = Right(strTxt2, Len(strTxt2) - 1)
        Loop
        
        If InStr(1, strBig, strTxt1) > 0 And Replace(strTxt1, vbNewLine, "") <> "" Then
            strCommon = strTxt2
            Exit For
        End If
    Next
    
    If UBound(Split(strCommon, vbNewLine)) < 3 Then Exit Function
    GetCommonMid = strCommon
End Function

Private Sub GetBeforeAndAfter(arrTxt As Variant, arrMid As Variant, ByRef strBefore As String, ByRef strAfter As String)
    Dim i As Long, j As Long, blnTmp As Boolean
    
    For i = 0 To UBound(arrTxt)
        If UCase(TrimEx(arrTxt(i))) = UCase(TrimEx(arrMid(0))) Then
            For j = 1 To UBound(arrMid)
                If UCase(TrimEx(arrTxt(i + j))) = UCase(TrimEx(arrMid(j))) Then
                    blnTmp = True
                Else
                    blnTmp = False
                    Exit For
                End If
            Next
        End If
        
        If j = UBound(arrMid) + 1 Or blnTmp Then
            Exit For
        End If
    Next
    
    strBefore = GetStrFromArr(arrTxt, 0, i - 1)
    strAfter = GetStrFromArr(arrTxt, i + UBound(arrMid) + 1, UBound(arrTxt))
End Sub


Private Sub CheckDiff(ByRef strTxt1 As String, ByRef strTxt2 As String)
    '�Ա��ı�����,�����и�ʽ��
    Dim arrTxt1 As Variant, arrTxt2 As Variant, arrMaxtrix() As Integer
    Dim lngFri As Long, lngSec As Long
    Dim i As Long, j As Long, lngS As Long
    
    If strTxt1 <> "" Then
        arrTxt1 = Split(strTxt1, vbNewLine)
        lngFri = UBound(arrTxt1) + 1
    End If
    If strTxt2 <> "" Then
        arrTxt2 = Split(strTxt2, vbNewLine)
        lngSec = UBound(arrTxt2) + 1
    End If
    
    If strTxt1 <> "" And strTxt2 <> "" Then
        arrMaxtrix = CreateMatrix(arrTxt1, arrTxt2)
    End If
    i = lngFri: j = lngSec
    strTxt1 = "": strTxt2 = ""
    Do While i <> 0 And j <> 0
        If arrMaxtrix(i, j) = 1 Then
            strTxt1 = arrTxt1(i - 1) & vbNewLine & strTxt1
            strTxt2 = arrTxt2(j - 1) & vbNewLine & strTxt2
            i = i - 1
            j = j - 1
        ElseIf arrMaxtrix(i, j) = 3 Then
            strTxt1 = vbNewLine & strTxt1
            strTxt2 = arrTxt2(j - 1) & vbNewLine & strTxt2
            j = j - 1
        Else
            strTxt1 = arrTxt1(i - 1) & vbNewLine & strTxt1
            strTxt2 = vbNewLine & strTxt2
            i = i - 1
        End If
    Loop
    
    '�������鲻ͬ��,��Ҫ�ٴ�ѭ��
    Do While i <> 0
        strTxt1 = arrTxt1(i - 1) & vbNewLine & strTxt1
        strTxt2 = vbNewLine & strTxt2
        i = i - 1
    Loop
    Do While j <> 0
        strTxt1 = vbNewLine & strTxt1
        strTxt2 = arrTxt2(j - 1) & vbNewLine & strTxt2
        j = j - 1
    Loop
End Sub

Private Function CreateMatrix(strFriT As Variant, strSecT As Variant) As Variant
    '����:���ݴ�������ݽ�������
    'strFriT-��һ������,strSecT-�ڶ�������
    Dim arrResult() As Integer, intMatrix() As Variant
    Dim lngFri As Long, lngSec As Long
    Dim i As Long, j As Long
    
    lngFri = UBound(strFriT) + 1: lngSec = UBound(strSecT) + 1
    ReDim intMatrix(lngFri, lngSec)
    ReDim arrResult(lngFri, lngSec)
    
    For i = 0 To lngFri
        intMatrix(i, 0) = 0
    Next
    For j = 0 To lngSec
        intMatrix(0, j) = 0
    Next
    
    'ѭ������,���и�ֵ
    '1-�I 2-�� 3-�� ,��ʾ��ȡ��̬�滮��·��,��Ϊ�ַ����Ƚ��� ���������ִ���
    For i = 1 To lngFri
        For j = 1 To lngSec
            If ConvertStr(strFriT(i - 1)) = ConvertStr(strSecT(j - 1)) Then
                intMatrix(i, j) = intMatrix(i - 1, j - 1) + 1
                arrResult(i, j) = 1
            ElseIf intMatrix(i - 1, j) >= intMatrix(i, j - 1) Then
                intMatrix(i, j) = intMatrix(i - 1, j)
                arrResult(i, j) = 2
            Else
                intMatrix(i, j) = intMatrix(i, j - 1)
                arrResult(i, j) = 3
            End If
        Next
    Next
    
    CreateMatrix = arrResult
    Exit Function
End Function


Private Sub DelElmentByIdx(ByRef arrElement As Variant, ByVal lngIdx As Long)
    '����������±�ɾ��Ԫ��
    Dim i As Long
    
    For i = lngIdx To UBound(arrElement) - 1
        arrElement(i) = arrElement(i + 1)
    Next
    
    ReDim Preserve arrElement(UBound(arrElement) - 1)
End Sub



Public Sub MergeDiffInto2SynEdit(ByVal strTxt1 As String, ByVal strTxt2 As String, txt1 As SyntaxEdit, txt2 As SyntaxEdit, colDiff As Collection)
    '�Ա��ı�����,������������������2���ؼ���
    Dim arrTxt1() As String, arrTxt2() As String
    Dim i As Long
    
    Set colDiff = New Collection
    txt1.Text = "": txt2.Text = ""
    arrTxt1 = Split(strTxt1, vbNewLine)
    arrTxt2 = Split(strTxt2, vbNewLine)
    
    For i = 0 To UBound(arrTxt1)
        txt1.RowText(i + 1) = arrTxt1(i)
        txt2.RowText(i + 1) = arrTxt2(i)
        txt1.SetRowColor i + 1, ��ɫ
        txt2.SetRowColor i + 1, ��ɫ
        
        If ConvertStr(arrTxt1(i)) <> ConvertStr(arrTxt2(i)) Then
            
            If arrTxt1(i) = "" Then '����
                txt2.SetRowColor i + 1, ��ɫ
                colDiff.Add arrTxt2(i), "_" & i + 1
            ElseIf arrTxt2(i) = "" Then
                txt1.SetRowColor i + 1, ��ɫ
                colDiff.Add arrTxt1(i), "_" & i + 1
            Else
                txt1.SetRowColor i + 1, ��ɫ
                txt2.SetRowColor i + 1, ��ɫ
                colDiff.Add arrTxt2(i), "_" & i + 1
            End If
        End If
        
    Next
End Sub

Public Sub MergeDiffInto1SynEdit(ByVal strTxt1 As String, ByVal strTxt2 As String, ByVal txtEdit As SyntaxEdit, ByRef colDiff As Collection, Optional ByVal blnDisplay = True)
    '�Ա��ı�����,������������������1���ؼ���
    Dim arrTxt1() As String, arrTxt2() As String
    Dim i As Long
    
    txtEdit.Text = ""
    Set colDiff = New Collection
    arrTxt1 = Split(strTxt1, vbNewLine)
    arrTxt2 = Split(strTxt2, vbNewLine)
    
    '�����ֺ���ɫ���Ƶ��ı��ؼ�
    For i = 0 To UBound(arrTxt1)
        With txtEdit
            .SetRowColor i + 1, ��ɫ
            If ConvertStr(arrTxt1(i)) <> ConvertStr(arrTxt2(i)) Then
                
                If arrTxt1(i) = "" Then '�����ı�
                    colDiff.Add "�տ�", "_" & (i + 1)
                    If blnDisplay Then
                        .SetRowColor i + 1, ��ɫ
                    End If
                    .RowText(i + 1) = arrTxt2(i)
                ElseIf arrTxt2(i) = "" Then 'ɾ������
                    If blnDisplay Then
                        .RowText(i + 1) = arrTxt1(i)
                        .SetRowColor i + 1, ��ɫ
                        colDiff.Add "�տ�", "_" & (i + 1)
                    Else
                        colDiff.Add arrTxt1(i), "_" & (i + 1)
                    End If
                Else
                    colDiff.Add arrTxt1(i), "_" & (i + 1)
                    .RowText(i + 1) = arrTxt2(i)
                    If blnDisplay Then
                        .SetRowColor i + 1, ��ɫ
                    End If
                End If
            Else
                txtEdit.RowText(i + 1) = arrTxt1(i)
            End If
        End With
    Next

End Sub

Public Function GetDiffRow(ByVal intType As Integer, ByVal lngCurRow As Long, ByRef arrPosition As Variant) As Long
    '���Ҳ����к�
    'intType-�������� 1-��һ��, 2-��һ��  lngCurRow:��ǰ�� arrPosition-λ������
    Dim lngResult As Long
    Dim i As Long, lngLast As Long, lngFirst As Long

    lngFirst = 1
    lngLast = UBound(arrPosition)

    If intType = 1 Then
        If lngCurRow < arrPosition(lngFirst) Then GetDiffRow = arrPosition(lngFirst): Exit Function
        If lngCurRow >= arrPosition(lngLast) Then GetDiffRow = arrPosition(lngFirst): Exit Function     '�����,�ص���һ��
    Else
        If lngCurRow <= arrPosition(lngFirst) Then GetDiffRow = arrPosition(lngLast): Exit Function
        If lngCurRow > arrPosition(lngLast) Then GetDiffRow = arrPosition(lngLast): Exit Function   '����С,�ص����һ��
    End If
    
    Do While lngFirst <> lngLast - 1
        If lngCurRow >= arrPosition(lngFirst) And lngCurRow < arrPosition((lngLast + lngFirst) \ 2) Then    '���ַ�
            lngLast = (lngLast + lngFirst) \ 2
        Else
           lngFirst = (lngLast + lngFirst) \ 2
        End If
    Loop
    
    '�������ͷ���ֵ
    If intType = 1 Then
        If arrPosition(lngLast) < lngCurRow Then
            lngResult = arrPosition(lngLast + 1)
        Else
            lngResult = arrPosition(lngLast)
        End If
    Else
        If arrPosition(lngFirst) < lngCurRow Then
            lngResult = arrPosition(lngFirst)
        Else
            lngResult = arrPosition(lngFirst - 1)
        End If
    End If
    GetDiffRow = lngResult
End Function


Public Function GetValueFromCol(ByRef colTxt As Collection, ByVal strIdx As String, Optional strErr As String) As String
    '����:����ָ���������Ӽ����л�ȡֵ,����������󷵻ؿ�
    Dim strResult As String
    
    On Error Resume Next
    strErr = ""
    strResult = colTxt.Item(strIdx)
    
    If err.Number <> 0 Then
        strErr = err.Description
    Else
        GetValueFromCol = strResult
    End If
    
End Function
