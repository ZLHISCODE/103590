Attribute VB_Name = "mEPREditor"
Option Explicit
'�жϵ�ǰ�Ƿ�ĳ����������»��߷ſ�

'################################################################################################################
'## ���ܣ�  ���������ı�����ָ���ؼ�������Ķ�λ��Ϣ
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey           :   IN  �����������ҵĹؼ���ID�š�
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ����ҵ��ùؼ��־���λ�ã��򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  �жϸ���λ���Ƿ����κ�һ���ؼ��ֶ�֮�䣬����ǣ������ؼ������λ�ú�ID��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         lngCurPosition  :   OUT ��ָ���ĵ�ǰλ��
'##         strKeyType      :   OUT �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey          :   OUT �����������ҵĹؼ���Key��
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ���������ĳ���ؼ��ֶ�֮�䣬�򷵻�True�����򷵻�False
'################################################################################################################
Public Function IsBetweenAnyKeys(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean

    '����������ʹ�� Instr() �� InstrRev() ���в��ң�
    Dim N As Long, i As Long, j As Long, k As Long
    Dim lFirst As Long
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    strKeyType = ""
    lngKSS = 0
    lngKSE = 0
    lngKES = 0
    lngKEE = 0
    lngKey = 0
    blnNeeded = False

    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        For N = 1 To UBound(gKeyWords)     '�� 5 �Ա����ؼ���
            '���Ƿ��ǹؼ���
            i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
            i = InStr(i, sText, gKeyWords(N).KeyEnd)    '�����������β�ؼ���
            If i <> 0 Then
                If .Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                    i = i + 1
                    GoTo LL1
                End If
                j = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL2:
                j = InStr(j, sText, gKeyWords(N).KeyStart) '���ҵ���β�ؼ��֣�����ͬ���Ŀ�ʼ�ؼ���
                If j <> 0 Then
                    If .Range(j - 1, j).Font.Hidden = False Then
                        j = j + 1
                        GoTo LL2
                    End If
                End If
                If (j = 0) Or (j > 0 And i < j) Then '���ڹؼ��ֶ�֮��
                    k = lngCurPosition
LL3:
                    k = InStrRev(sText, gKeyWords(N).KeyStart, k, vbTextCompare)     '��ƥ��Ŀ�ʼ�ؼ���
                    If k <> 0 Then
                        If .Range(k - 1, k).Font.Hidden = False Then
                            k = k - 1
                            GoTo LL3
                        End If
                        strKeyType = Left(gKeyWords(N).KeyStart, 1)
                        lngKSS = k - 1
                        lngKSE = k + 15
                        lngKES = i - 1
                        lngKEE = i + 15
                        lngKey = Val(.Range(k + 2, k + 10).Text)
                        blnNeeded = -Val(.Range(k + 11, k + 12).Text)
                        IsBetweenAnyKeys = True
                        Exit For
                    End If
                End If
            End If
        Next N
    End With
End Function

'################################################################################################################
'## ���ܣ�  �жϸ���λ���Ƿ���ָ���Ĺؼ��ֶ�֮�䣬����ǣ������ؼ������λ�ú�ID��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         lngCurPosition  :   IN  ��ָ���ĵ�ǰλ�ã���1��ʼ��ţ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKey          :   OUT �����������ҵĹؼ���Key��
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ���������ָ���ؼ��ֶ�֮�䣬�򷵻�True�����򷵻�False
'################################################################################################################
Public Function IsBetweenKeys(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean

    Dim N As Long, i As Long, j As Long, k As Long
    Dim lFirst As Long
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    lngKSS = 0
    lngKSE = 0
    lngKES = 0
    lngKEE = 0
    lngKey = 0
    blnNeeded = False

    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        
        '���Ƿ��ǹؼ���
        If lngCurPosition = 0 Then lngCurPosition = edtThis.SelStart + 1
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, strKeyType & "E")    '�����������β�ؼ���
        If i <> 0 Then
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            j = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL2:
            j = InStr(j, sText, strKeyType & "S") '���ҵ���β�ؼ��֣�����ͬ���Ŀ�ʼ�ؼ���
            If j <> 0 Then
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
            End If
            If (j = 0) Or (j > 0 And i < j) Then '���ڹؼ��ֶ�֮��
                k = lngCurPosition
LL3:
                k = InStrRev(sText, strKeyType & "S", k, vbTextCompare)    '��ƥ��Ŀ�ʼ�ؼ���
                If k <> 0 Then
                    If .TOM.TextDocument.Range(k - 1, k).Font.Hidden = False Then
                        k = k - 1
                        GoTo LL3
                    End If
                    lngKSS = k - 1
                    lngKSE = k + 15
                    lngKES = i - 1
                    lngKEE = i + 15
                    lngKey = Val(.TOM.TextDocument.Range(k + 2, k + 10))
                    blnNeeded = -Val(.TOM.TextDocument.Range(k + 11, k + 12))
                    IsBetweenKeys = True
                End If
            End If
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  ����ָ��λ�ú�ĵ�һ���������͵Ĺؼ���λ��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         lngCurPosition  :   IN  ��ָ���ĵ�ǰλ�ã���1��ʼ��ţ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey          :   OUT �����������ҵĹؼ���Key��
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ���������ָ���ؼ��ֶ�֮�䣬�򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  ����ָ��λ��ǰ�ĵ�һ���������͵Ĺؼ���λ��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         lngCurPosition  :   IN  ��ָ���ĵ�ǰλ�ã���1��ʼ��ţ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey          :   OUT �����������ҵĹؼ���Key��
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ���������ָ���ؼ��ֶ�֮�䣬�򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindPrevKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStrRev(sText, sTMP, i)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindPrevKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ����λ�õ���һ������ؼ���λ�ã�������֣������ؼ������λ�ú�ID��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         lngCurPosition  :   IN  ��ָ���ĵ�ǰλ�ã���1��ʼ��ţ�
'##         strKeyType      :   OUT �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey          :   OUT �����������ҵĹؼ���Key��
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ���������ָ���ؼ��ֶ�֮�䣬�򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindNextAnyKey(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 14
                lngKES = j - 2
                lngKEE = j + 14
                lngKey = Val(.TOM.TextDocument.Range(i + 1, i + 9))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 10, i + 11))
                FindNextAnyKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  ����ָ���ı��εĳ�����ʽ�����������ʽ�Ͷ����ʽ��
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         varStyle        :   IN  ��ָ������ʽ�����Ϊ���֣����ʾ������ʽ��ţ�������ַ��������ʾ������ʽ���ƣ�
'##         lStart          :   IN  ����Ҫ������ʽ���ı���ʼλ��
'##         lEnd            :   IN  ����Ҫ������ʽ���ı�����λ��
'##         ForceParaFmt    :   IN  ���Ƿ�ǿ������ָ����Χ���ı��Ķ�������
'##
'## ˵����  ͨ��ָ��λ��λ�ڶ���ʱ��ͬʱ�����������ԺͶ������ԣ�����Ƕ����ı�����ֻ�����������ԣ�
'##         �����Ҫͬʱ���ö����ı��������ԺͶ������ԣ������ ForceParaFmt Ӧ����Ϊ True��
'##
'##         �� varStyle =0 ���ߡ������ʽ������ͬʱ��������ʽ��ForceParaFmt = Trueʱ���������ʽ
'##
'##   ��  ���ĳ������Ϊ tomUndefined = -9999999 ����ʾ���ı���������ֵ�����磬�����м��Ϊ -9999999����ʾ���ı��м�ࡣ
'################################################################################################################
Public Sub SetCommonStyle(ByRef edtThis As Object, _
        ByVal varStyle As Variant, _
        ByVal lStart As Long, _
        ByVal lEnd As Long, _
        Optional ForceParaFmt As Boolean = False)
        
    Dim rs As New ADODB.Recordset, blnForceEdit As Boolean
    Dim strFont As String, strPara As String
    Dim T As Variant
    Dim blnBeginWithCRLF As Boolean
    
    blnForceEdit = edtThis.ForceEdit
    
    If IsNumeric(varStyle) Then
        If varStyle = 0 Then
            If ForceParaFmt Then edtThis.TOM.TextDocument.Selection.Para.Reset tomDefault
            edtThis.TOM.TextDocument.Selection.Font.Reset tomDefault
            Exit Sub
        End If
        gstrSQL = "select * from ����������ʽ where ���=[1]"
    Else
        If varStyle = "�����ʽ" Then
            If ForceParaFmt Then edtThis.TOM.TextDocument.Selection.Para.Reset tomDefault
            edtThis.TOM.TextDocument.Selection.Font.Reset tomDefault
            Exit Sub
        End If
        gstrSQL = "select * from ����������ʽ where ����=[1]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, varStyle)
    
    If Not rs.EOF Then
        strFont = rs("������ʽ")
        strPara = rs("������ʽ")
        
        If edtThis.Range(lStart - 2, lStart) = vbCrLf Or edtThis.SelStart = 0 Then
            blnBeginWithCRLF = True
        Else
            blnBeginWithCRLF = False
        End If
        
        T = Split(strFont, ";")
        If UBound(T) > 0 Then
            With edtThis.Range(lStart, lEnd).Font
                edtThis.ForceEdit = True
                If Trim(T(0)) <> "" Then .Name = T(0)
                If Val(T(1)) > 0 Then .Size = Val(T(1))
                
                .Bold = IIf(Mid(T(2), 1, 1) = 1, True, False)
                .Italic = IIf(Mid(T(2), 2, 1) = 1, True, False)
'               .Hidden = IIf(Mid(T(2), 3, 1) = 1, True, False)
'                .Protected = IIf(Mid(T(2), 4, 1) = 1, True, False)
'                .Link = IIf(Mid(T(2), 5, 1) = 1, True, False)
'                .Strikethrough = IIf(Mid(T(2), 6, 1) = 1, True, False)
                .Superscript = IIf(Mid(T(2), 7, 1) = 1, True, False)
                .Subscript = IIf(Mid(T(2), 8, 1) = 1, True, False)
                
'                .Underline = Val(T(3))
'                .BackColor = Val(T(4))
'                .ForeColor = Val(T(5))
                edtThis.ForceEdit = blnForceEdit
            End With
        End If

        T = Split(strPara, ";")
        If UBound(T) > 0 Then
            If blnBeginWithCRLF Or ForceParaFmt Then
            '���ö�����ʽ
                With edtThis.Range(lStart, lEnd).Para
                    edtThis.ForceEdit = True
                    '���Ϊ9���򲻸ı�ֵ
                    If Mid(T(0), 2, 1) < 9 Then .ListAlignment = Mid(T(0), 2, 1)                       '����ȡֵΪ��0��1��2
'                   If Mid(T(0), 3, 1) < 9 Then .LineSpacingRule = IIf(Mid(T(0), 3, 1) = 1, True, False) '����ȡֵΪ��0��1��2��3��4��5
                    
                    If Val(T(1)) <> -9999999 Then .Style = Val(T(1))        '����ȡֵ�� -1 ~ -10
                    If Val(T(2)) <> -9999999 Then
                        .ListType = Val(T(2))     '����ȡֵ��0 �� 6��65536��131072��196608
                        .ListStart = Val(T(3))
                    End If
                    If Val(T(4)) <> tomUndefined Then .FirstLineIndent = Val(T(4)) '��������һ��������
                    If Val(T(5)) <> tomUndefined Then .LeftIndent = Val(T(5))
                    If Val(T(6)) <> tomUndefined Then .RightIndent = Val(T(6))
'                   If Val(T(7)) <> tomUndefined Then .LineSpacing = Val(T(7))
                    If Val(T(8)) <> tomUndefined Then .ListTab = Val(T(8))
                    If Val(T(9)) <> tomUndefined Then .SpaceBefore = Val(T(9))
                    If Val(T(10)) <> tomUndefined Then .SpaceAfter = Val(T(10))
                    
                    If Mid(T(0), 3, 1) < 9 And Val(T(7)) <> tomUndefined Then .SetLineSpacing Mid(T(0), 3, 1), Val(T(7))
                    If Mid(T(0), 1, 1) < 9 Then .Alignment = Mid(T(0), 1, 1)                           '����ȡֵΪ��0��1��2
                    edtThis.ForceEdit = blnForceEdit
                End With
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

