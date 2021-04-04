Attribute VB_Name = "mdlTransfusion"
Option Explicit
Public gblnShowInTaskBar As Boolean         '�Ƿ���ʾ��������������

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String            'OEM��Ʒ����
Public glngSys As Long                      'ϵͳ���
Public glngModul As Long                    'ģ���
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
'ϵͳ����
Public gbytCardLen As Byte '���￨�ų���
'Public gblnCardHide As Boolean '���￨��������ʾ

Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gint�Һ����� As Integer '�Һŵ���Ч����
Public gbln�������Ҷ��� As Boolean
Public gint�����Դ As Integer '1-��ҽ��ѡ��������Դ,2-������ϱ�׼����,3-���ռ�����������
Public gint������� As Integer '1-������������,2-�����ݿ���ȡ����,3-��ҽ�����˴����ݿ�����
Public gblnִ�к���� As Boolean    'ִ�к��Զ���˻��۵�
Public gbln������֤ As Boolean '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
Public gobjPlugIn As Object
Public gstrҩƷ�۸�ȼ� As String 'Ժ����ҩƷ�۸�ȼ�
Public gstr���ļ۸�ȼ� As String 'Ժ�������ļ۸�ȼ�
Public gstr��ͨ��Ŀ�۸�ȼ� As String 'Ժ������ͨ��Ŀ�۸�ȼ�

Public gstrҽ���˶� As String    '��ѪƤ��ҽ����Ҫ�˶� ��λ��ȡ11����һλΪ ��Ѫҽ�����ڶ�λΪ Ƥ��ҽ��

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum enuCardProperty
    ���� = 0
    ȫ�� = 1
    �ɶ��� = 2
    �����ID = 3
    ���ų��� = 4
    ȱʡ��� = 5
    �����ʻ� = 6
    ����������ʾ = 7
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = zlCommFun.NVL(rsTmp!����ID, 0)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            GetUserInfo = True
        End If
    End If
End Function

Public Function GetSquareCardInfo(ByVal strSquareCards As String, ByVal strCardName As String, ByVal intElement As Integer) As String
'���ܣ�ȡһ��ͨ���ŵĳ���
'������
'  strSquareCards��һ��ͨ������Ϣ
'  strCardName��ָ����ȡ�Ŀ�������
'  intElement��ָ��ȡһ��ͨ��Ϣ����Ԫ��
'���أ����ų���
    
    If strSquareCards = "" Then Exit Function
    
    Dim i As Integer
    Dim arrInfo As Variant
    Dim strTmp As String
    
    GetSquareCardInfo = ""
    
    On Error GoTo errHandle
    arrInfo = Split(strSquareCards, ";")
    For i = LBound(arrInfo) To UBound(arrInfo)
        strTmp = Split(arrInfo(i), "|")(enuCardProperty.ȫ��)
        If strCardName = strTmp Then
            GetSquareCardInfo = Split(arrInfo(i), "|")(intElement)
            Exit For
        End If
    Next
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPara As String
    
    On Error GoTo errH
        
    '���￨����ĳ���
    gbytCardLen = 7
    strSQL = "select ���ų��� from ҽ�ƿ���� where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���￨����", "���￨")
    If Not rsTmp.EOF Then
        gbytCardLen = IIf(IsNull(rsTmp!���ų���), 7, rsTmp!���ų���)
    End If
    
    'HISϵͳ����
    
    '�Һ���Ч����
    strPara = zlDatabase.GetPara(21, glngSys)
    gint�Һ����� = zlCommFun.NVL(strPara, 0)
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = zlCommFun.NVL(zlDatabase.GetPara(23, glngSys), 0)
    
    '���������Դ
    gint�����Դ = zlCommFun.NVL(zlDatabase.GetPara(55, glngSys), 1)
    
    '������뷽ʽ
    gint������� = zlCommFun.NVL(zlDatabase.GetPara(65, glngSys), 1)
    '�����Ϳ����Ƿ��������
    gbln�������Ҷ��� = Val(zlDatabase.GetPara(99, glngSys)) <> 0
    'һ��ͨ������֤
    gbln������֤ = Val(zlDatabase.GetPara(28, glngSys)) <> 0 '���ﲡ������ʱ��Ҫˢ����֤
    
    '��Ŀִ��ǰ�������շѻ��ȼ������
    
    '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
    gstrҽ���˶� = zlDatabase.GetPara(186, glngSys)
    
    'ִ�к��Զ����
    gblnִ�к���� = Val(zlDatabase.GetPara(81, glngSys)) <> 0 'ִ�к��Զ���˻��۵�
    Call InitPriceLevel
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
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

Public Function CacleTransTime(ByVal Һ������ As Long, ByVal ��ϵ�� As Long, ByVal ÿ���ӵ��� As Integer) As Integer
    '������Һʱ��
    '��Һʱ��(����)=(Һ������(ml)����ϵ��)/(ÿ���ӵ���)
    If ÿ���ӵ��� > 0 Then
        CacleTransTime = (Һ������ * ��ϵ��) / ÿ���ӵ���
    End If
End Function

Public Function GetAdvicePause(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡָ��ҽ������ͣʱ��μ�¼
'���أ�"��ͣʱ��,��ʼʱ��;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select ��������,����ʱ�� From ����ҽ��״̬" & _
        " Where �������� IN('6','7') And ҽ��ID=[1]" & _
        " Order by ����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = "6" Then
            strTmp = strTmp & ";" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!�������� = "7" Then
            '���õ���һ�벻����ͣ�ķ�Χ֮��
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ�������Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
'˵��������ʱ���ж�,����ͣ���ڰ���ʼ����ֹ�����ж�
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '������δ���û���ͣ��ʱ��ֹͣ
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function Between(X, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function Calc�����ڿ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datĳ��ִ��ʱ�� As Date, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Date
'���ܣ����ݳ�����ĳ��ִ��ʱ�䣬�õ����ڸ������ڵĿ�ʼ��׼ʱ��
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat��ʼִ��ʱ��
    datBegin = datCurr
    If str�����λ = "��" Then datCurr = Format(datCurr - (Weekday(datCurr, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
    
    Do While datCurr <= datĳ��ִ��ʱ��
        datBegin = datCurr
        If str�����λ = "��" Then
            datCurr = datCurr + 7
        ElseIf str�����λ = "��" Then
            datCurr = datCurr + intƵ�ʼ��
        ElseIf str�����λ = "Сʱ" Then
            datCurr = DateAdd("h", intƵ�ʼ��, datCurr)
        End If
    Loop
    Calc�����ڿ�ʼʱ�� = datBegin
End Function

Public Function Calc���ڷֽ�ʱ��(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ���ʱ��μ�����εķֽ�ִ��ʱ�估����
'������datBegin-datEnd=Ҫ�����ʱ���,����datBeginӦΪÿ�����ڵĿ�ʼ��׼ʱ��
'      strPause=��ͣ��ʱ���
'      dat��������=��������ʱ��������
'���أ�"ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss),ʱ�������Ϊ����
'˵����1.ʱ�����Ҫ�ų���ͣ��ʱ���,����������˶�����
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If
        
    vCurTime = datBegin
    
    If str�����λ = "��" Then
        vCurTime = zlCommFun.GetWeekBase(datBegin)
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = zlCommFun.GetWeekBase(dat��������))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
                        
            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                If i - 1 <= UBound(arrTime) Then '���ܿ��ܴ�������
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "��" Then
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat��������))
        Else
            blnFirst = False
        End If
        
        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
            
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + intƵ�ʼ��, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= Format(datBegin, "yyyy-MM-dd HH:mm:ss") And vTmpTime <= Format(datEnd, "yyyy-MM-dd HH:mm:ss") Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime
            
            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    
    Calc���ڷֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '������δ���û���ͣ��ʱ��ֹͣ
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

'Public Function GetOwner(ByVal lngSys As Long) As String
''���ܣ���ȡָ��ϵͳ��������
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL  As String
'
'    On Error GoTo errH
'    strSQL = "Select ������ From zlSystems Where ���=[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetOwner", lngSys)
'    If Not rsTmp.EOF Then
'        GetOwner = rsTmp!������
'    End If
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
    '������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = zlCommFun.NVL(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01"), "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function DelInvalidChar(ByVal strChar As String, Optional ByVal strInvalidChar As String) As String
    'ɾ���Ƿ��ַ�
    'strChar: Ҫ������ַ�
    'strInvalidChar���Ƿ��ַ��������Ϊ�գ���Ϊ~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,���򰴴�����ַ�����
    Dim strBit As String, i As Integer, strWord As String
    strWord = "~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>"
    If strInvalidChar <> "" Then strWord = strInvalidChar
    If Len(strChar) > 0 Then
        For i = 1 To Len(strChar)
            strBit = Mid$(strChar, i, 1)
            If InStr(strWord, strBit) <= 0 Then
                DelInvalidChar = DelInvalidChar & strBit
            End If
        Next
    End If
End Function

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
End Function

'Public Function GetAdviceMoney(ByVal str��ID As String, ByVal strҽ��ID As String, ByVal str���ͺ� As String, _
'    str��� As String, str����� As String, ByVal bln����ִ�� As Boolean, ByVal byt��Դ As Byte) As Currency
''���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
''������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
''      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
''      byt��Դ��1:���2-סԺ
''���أ�str���,str�����=���ڱ�����ʾ
''˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL As String, curMoney As Currency
'    Dim strTab As String
'
'    str��� = "": str����� = ""
'
'    On Error GoTo errH
'
'    If zldatabase.GetPara(81, glngSys) <> "1" Then Exit Function
'    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
'
'    If bln����ִ�� Then
'        strSQL = _
'            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
'            " From " & strTab & " A,�շ���Ŀ��� B" & _
'            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
'            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
'            "       Union All" & _
'            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
'            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
'            " Group by B.����,B.����"
'    Else
'        strSQL = _
'            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
'            " From " & strTab & " A,�շ���Ŀ��� B" & _
'            " Where A.ҽ����� + 0 In" & _
'            "      (Select ID From ����ҽ����¼" & _
'            "       Where ID In (Select Column_Value From Table(f_Num2list([1])))" & _
'            "       Union All" & _
'            "       Select ID From ����ҽ����¼" & _
'            "       Where ���id In (Select Column_Value From Table(f_Num2list([1]))))" & _
'            "  And (A.��¼����, A.NO) In" & _
'            "      (Select ��¼����, NO From ����ҽ������" & _
'            "       Where ҽ��id In" & _
'                "      (Select ID From ����ҽ����¼" & _
'                "       Where ID In (Select Column_Value From Table(f_Num2list([1])))" & _
'                "       Union All" & _
'                "       Select ID From ����ҽ����¼" & _
'                "       Where ���id In (Select Column_Value From Table(f_Num2list([1]))))" & _
'            "         And ���ͺ� + 0 In (Select Column_Value From Table(f_Num2list([3])))" & _
'            "       Union All" & _
'            "       Select ��¼����, NO From ����ҽ������" & _
'            "       Where ҽ��id In (Select Column_Value From Table(f_Num2list([2])))" & _
'            "         And ���ͺ� + 0 In (Select Column_Value From Table(f_Num2list([3]))))" & _
'            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
'            " Group by B.����,B.����"
'    End If
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str��ID, strҽ��ID, str���ͺ�, glngSys)
'
'    curMoney = 0
'    Do While Not rsTmp.EOF
'        curMoney = curMoney + Val("" & rsTmp!���)
'        str��� = str��� & rsTmp!����
'        str����� = str����� & "," & rsTmp!����
'        rsTmp.MoveNext
'    Loop
'
'    str����� = Mid(str�����, 2)
'    GetAdviceMoney = curMoney
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function OneCardCheck(ByVal lngҽ��ID_IN As Long, ByVal lng���ͺ�_IN As Long, _
                             Optional frmMain As Object, Optional objCardSquare As Object) As Integer
    'һ��ͨ������
    '����: 0- ���������ߣ�1-���������ߣ��ɹ���2-���������ߣ�ʧ��
    
    Dim lng����ID As Long, curMoney As Currency, strSQL As String
    Dim str��� As String, str����� As String, strNO As String, lng��¼���� As Long
    Dim rsTmp As ADODB.Recordset
    On Error GoTo hErr
    OneCardCheck = 2    'Ĭ��������ʧ��
    
    strSQL = "Select ����ID,������� From ����ҽ����¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ҽ��ִ�����", lngҽ��ID_IN)
    If Not rsTmp.EOF Then
        lng����ID = Val("" & rsTmp!����ID)
        str��� = Trim("" & rsTmp!�������)
    End If
    
    strSQL = "Select No,��¼���� From ����ҽ������ Where ҽ��id=[1] And ���ͺ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ҽ��ִ�����", lngҽ��ID_IN, lng���ͺ�_IN)
    If Not rsTmp.EOF Then
        strNO = Trim("" & rsTmp!NO)
        lng��¼���� = Val("" & rsTmp!��¼����)
    End If
    
    If zlDatabase.GetPara("��Ŀִ��ǰ�������շѻ��ȼ������", glngSys) = 1 Then
        If objCardSquare Is Nothing Then
            'һ��ͨ���Ѳ���δ�����ɹ���
            MsgBox "һ��ͨ���Ѳ���δ�����ɹ���", vbQuestion, frmMain.Caption
            Exit Function
        End If
        '�µ�һ��ͨ����
        
        '1-����ģʽ
        If lng��¼���� = 2 Then
            If gblnִ�к���� Then
                'ԭ���Ĺ��ܴ���
                OneCardCheck = 0
            Else
                '1.ˢ����ȡ����
                '2.�Ƿ����δ��˵Ļ��۵�
                'If ItemHaveCash(1, False, lngҽ��ID_IN, lngҽ��ID_IN, lng���ͺ�_IN, str���, strNO, lng��¼����, 0, 0) Then
                    '������� zlSquareAffirm
'                    frmMain Object  In  ������ö���
'                    lngModule   Long    IN  ���õ�ģ���
'                    strPrivs    String  In  Ȩ�޴�
'                    lngPatiID   Long    In  ����ID,���Բ���,�ڱ��ӿڴ�����ˢ��!
                     If Not objCardSquare.zlSquareAffirm(frmMain, glngModul, gstrPrivs, lng����ID, , False, , , lngҽ��ID_IN) Then
                        MsgBox "����ʧ�ܣ�����ִ�к���Ĳ�����", vbInformation, frmMain.Caption
                        Exit Function
                     End If
                'End If
            End If
        ElseIf lng��¼���� = 1 Then
            'ˢ����ȡ���ˣ��Ƿ����δ�շѵĻ��۵�
            'Dim strIDs As String
            'If ItemHaveCash(1, False, lngҽ��ID_IN, lngҽ��ID_IN, lng���ͺ�_IN, str���, strNO, lng��¼����, 0, 0, , , strIDs) Then

                If Not objCardSquare.zlSquareAffirm(frmMain, glngModul, gstrPrivs, lng����ID, , False, , , lngҽ��ID_IN) Then
                                                                                              
                    MsgBox "����ʧ�ܣ�����ִ�к���Ĳ�����", vbInformation, frmMain.Caption
                    Exit Function
                End If
            'End If
        End If
        OneCardCheck = 1    '�����̳ɹ�
    Else
        OneCardCheck = 0    '��Ŀִ��ǰ�������շѻ��ȼ������ ����δ������������
    End If
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Private Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
'    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
'    Optional ByVal blnMove As Boolean, Optional ByVal dat����ʱ�� As Date, Optional ByRef strҽ��IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
''���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
''������int������Դ=1-����,2-סԺ
''      str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
''      int��ʽ=0-����Ƿ����δ�շѼ�¼
''              1-����Ƿ�������շѼ�¼
''      int�������=1=סԺ���͵��������
''      ���أ�strҽ��IDs=��ҽ������ص�ҽ��ID,NOs=ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL As String, strTab As String
'
'    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
'        strTab = "סԺ���ü�¼"
'    Else
'        strTab = "������ü�¼"
'    End If
'    ItemHaveCash = True
'    strҽ��IDs = ""
'    strNOs = ""
'
'    '��Ӧ�ķ������Ƿ����δ�շ�[��������]������
'    '���嵥ֻ��ʾ���շѲ�ͬ��
'    '1.�����ҽ������(���Ӽ�¼���ʵ���������Ϊ���ܲ��շѵ�����ʵ�)
'    '2.���ʻ���Ҳ��ʾΪδ��(�嵥��Ҫ���Գ���ִ�к����)
'    '3.��NO��Ӧ�����ҽ���ķ��ü��(�嵥�ǰ���ʾ��ҽ��ID)
'    strSQL = _
'        " Select A.��¼״̬,Nvl(B.���ID,B.ID) as ҽ��ID,B.�������,A.ִ��״̬,A.NO" & _
'        " From " & strTab & " A,����ҽ����¼ B" & _
'        " Where A.NO=[4] And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And A.��¼����=[5]" & IIf(bln����ִ��, " And B.ID=[2]", "") & _
'        " Union ALL " & _
'        " Select B.��¼״̬,Nvl(C.���ID,C.ID) as ҽ��ID,C.�������,B.ִ��״̬,A.NO" & _
'        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
'        " Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0" & IIf(bln����ִ��, " And A.ҽ��ID=[2]", _
'            " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
'        " And A.���ͺ�=[3] And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����=[5]"
'    If blnMove Then
'        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
'        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
'        strSQL = Replace(strSQL, strTab, "H" & strTab)
'    ElseIf zldatabase.DateMoved(dat����ʱ��) Then
'        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
'    End If
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
'    If Not rsTmp.EOF Then
'        If int��ʽ = 0 Then
'            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ִ��״̬=9"
'            If Not rsTmp.EOF Then
'                blnIsAbnormal = True
'                ItemHaveCash = False
'            Else
'                rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬=0"
'                If Not rsTmp.EOF Then ItemHaveCash = False
'            End If
'
'            While Not rsTmp.EOF
'                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
'                    strҽ��IDs = strҽ��IDs & "," & rsTmp!ҽ��ID
'                End If
'                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
'                    strNOs = strNOs & "," & rsTmp!NO
'                End If
'                rsTmp.MoveNext
'            Wend
'            strNOs = Mid(strNOs, 2)
'            strҽ��IDs = Mid(strҽ��IDs, 2)
'        ElseIf int��ʽ = 1 Then
'            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬<>0 And ִ��״̬<>9"
'            If rsTmp.EOF Then ItemHaveCash = False
'        End If
'    ElseIf int��ʽ = 1 Then
'        ItemHaveCash = False
'    End If
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetMaxNoAddOne(ByVal strField As String, ByVal strTableAndWhere As String) As String
    '2012-06-04
    '��ȡָ����ָ���ֶε����ֵ��һ��һ�����ڳ�ʼ��ʱ�Զ������ı��
    Dim strSQL As String, rsTmp As ADODB.Recordset, strMaxNO As String
    On Error GoTo hErr
    strMaxNO = ""
    strSQL = "Select Max(" & strField & ") as MaxNo From " & strTableAndWhere
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxNoAddOne")
    Do Until rsTmp.EOF
        strMaxNO = Trim$("" & rsTmp!MaxNo)
        rsTmp.MoveNext
    Loop
    If strMaxNO <> "" Then
        If IsNumeric(strMaxNO) Then
            strMaxNO = Format(Val(strMaxNO) + 1, String(Len(strMaxNO), "0"))
        Else
            strMaxNO = zlCommFun.IncStr(strMaxNO)
        End If
    Else
        strMaxNO = "001"
    End If
    If strMaxNO <> "" Then GetMaxNoAddOne = strMaxNO
    Exit Function
hErr:
    GetMaxNoAddOne = ""
End Function

'Public Function GetDeptInListPara(ByVal strParaName As String, ByVal lngDeptID As Long) As Boolean
'    '��ȡ������Һ�еļ������ز���
''        ������Һ_��׼�����б�
''        ������Һ_���п����б�
''        ������Һ_�򵥴����б�
''        ������Һ_��Һ�����б�
''        ������Һ_Ѳ�ӿ����б�
'    Dim strTmp As String
'    strTmp = zlDatabase.GetPara(strParaName, glngSys)
'    If strTmp <> "" Then
'        If Left(strTmp, 1) <> "," Then strTmp = "," & strTmp
'        If Right(strTmp, 1) <> "," Then strTmp = strTmp & ","
'        GetDeptInListPara = InStr(strTmp, "," & lngDeptID & ",") > 0
'    Else
'        GetDeptInListPara = False
'    End If
'End Function

Public Function AllocationDesks(ByVal lngDeptID As Long, ByVal objPati As cPatient, _
        ByRef strSeqNo As String, ByRef strErr As String) As Boolean
    '��Һʱ ���䴩��̨
    '���룺
    '   lngDeptID  :����ID
    '   objPati    :������Ϣ����
    '������
    '   strSeqNo   :����Ĵ���̨���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSeqNo = "": strErr = ""
    strSQL = "Zl_���ﴩ��̨_Liquid(" & lngDeptID & "," & objPati.����ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "���䴩��̨")
    
    strSQL = "select ����̨, ״̬ From �ŶӼ�¼ Where ����ID=[1] And ����id=[2] order by ���� desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���䴩��̨", lngDeptID, objPati.����ID)

    If rsTmp.EOF = False Then
        If zlCommFun.NVL(rsTmp!����̨) = "" Then
            strErr = "û�д���̨���������ô���̨����ʹ�ô˹��ܣ�"
        Else
            strSeqNo = CStr(rsTmp!����̨)
            SaveOperLog lngDeptID, objPati, QUEUE, "���䵽" & strSeqNo & "�Ŵ���̨,��ǰ״̬Ϊ" & Trim$("" & rsTmp!״̬)
        End If
    End If
    
    If strSeqNo = "" Then
        strErr = "����ʧ�ܣ����Ժ����ԣ�"
    Else
        AllocationDesks = True
    End If
    Exit Function
    
hErr:
    AllocationDesks = False
    strErr = Err.Description
End Function

Public Function CurDayHaveItem(ByVal objPati As cPatient, ByVal lngDeptID As Long) As Boolean
    '�жϵ����Ƿ�ӹ���Һ�ĵ�
    '���أ���True �ӹ���Һ������alse��δ�ӹ���Һ��
    Dim objExe As New ExecRecord
    Dim dateS As Date, dateE As Date
    Dim i As Integer, Y As Integer
    Dim blnNoCall As Boolean
   
    Dim blnHaveItem As Boolean
    dateS = Format(objPati.�Һ�ʱ��, "yyyy-MM-dd 00:00:00")
    dateE = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    
    
    blnHaveItem = False
    Call objExe.GetExecGroups(objPati, lngDeptID, 1, dateS, dateE)
    
    dateS = Format(dateE, "yyyy-MM-dd 00:00:00")
    For i = 1 To objExe.Count - 1
        If objExe.Item(i).ִ�з��� = "1-��Һ" And objExe.Item(i).ִ��ʱ�� >= dateS And objExe.Item(i).ִ��ʱ�� <= dateE And objPati.����ʱ�� >= dateS And objPati.����ʱ�� <= dateE Then
            '����ӹ���Һ�ĵ����Ͳ��� ״̬��
            blnHaveItem = True
            Exit For
        End If
    Next
    
    CurDayHaveItem = blnHaveItem
    
End Function

Public Function CurDayNoCall(ByVal lngDeptID As Long, ByVal objPatients As cPatients, ByVal objCurPati As cPatient) As Boolean
    '2012-09-25 ��鵱���Ƿ��Ѿ��й����κŵ����
    '���أ�True �����Ѿ��ӹ���Һ����������Ҫ���У�����״̬����False������δ�ӹ���Һ������Ҫ���У���Ҫ���¸�״̬��
    Dim dateS As Date, dateE As Date
    Dim blnNoCall As Boolean, objPati As cPatient
    
    dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dateE = Format(dateS, "yyyy-MM-dd 23:59:59")
    blnNoCall = False
    For Each objPati In objPatients
        If objPati.����ID = objCurPati.����ID And _
            objPati.�Һŵ� <> objCurPati.�Һŵ� And _
            objPati.����ʱ�� >= dateS And objPati.����ʱ�� <= dateE And _
            (Val(objPati.�Ŷ�״̬) = 5 Or Val(objPati.�Ŷ�״̬) = 7) Then
            If Val(objPati.�Ŷ�״̬) = 5 Then
                blnNoCall = True
            Else
                blnNoCall = CurDayHaveItem(objPati, lngDeptID)
            End If
            SaveOperLog lngDeptID, objCurPati, QUEUE, "���˵����Ѿ���" & objPati.�Ŷ�״̬ & "��¼���Һŵ�Ϊ" & objPati.�Һŵ�
        End If
    Next
End Function

Public Function Liquid(ByVal lngDept As Long, ByVal strNO As String, ByVal objPatiList As cPatients, ByRef strErr As String) As String
'��Һ����
    
    Dim strSQL As String, strSeqNo As String
    Dim blnNoCall As Boolean
    Dim objPati As cPatient
    
    strErr = ""
    If strNO <> "" Then
        Err.Clear
        On Error Resume Next
        Set objPati = objPatiList.Item(strNO)
        If objPati Is Nothing Or Err.Number <> 0 Then
            '���뱾�߼����������Ŷ�״̬=4��������=3���˺ţ�=2�����š��������ݡ���ΪFetchPatientsֻȡ1��5��6��7����״̬���Ŷ�����
            Liquid = "5-������"
            Err.Clear
            SaveOperLog lngDept, strNO, QUEUE, "5-������"
            Exit Function
        End If
        
        strSeqNo = objPati.�Һŵ�
        If Err.Number <> 0 Then
            Liquid = "5-������"
            Err.Clear
            SaveOperLog lngDept, strSeqNo, QUEUE, "�Һŵ�"
            Exit Function
        End If
        
        blnNoCall = CurDayNoCall(lngDept, objPatiList, objPati)
        
        If Err.Number <> 0 Then
            Liquid = "5-������"
            SaveOperLog lngDept, strNO, QUEUE, "Call�쳣"
            Exit Function
        End If
        
        If Not blnNoCall Then
            '���䴩��̨������ǰ���ӵ��д��� 2012-10-10
'            If Not AllocationDesks(lngDept, strNo, strSeqNo, strErr) Then
'                Exit Function
'            End If
            Liquid = "5-������"
        Else
            Liquid = "7-ִ����"
        End If
    Else
        strErr = "��ѡ��һ����¼����ִ�д˲���!"
    End If

End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, strLabel As String, intResult As Integer)
'���ܣ���ȡƤ�Ա�ע�ͽ��
'������strScript=Ƥ�Խ������������"����(+),������(++);����(-)"
'      strSelect=��ѡ���Ƥ�Խ������������"����"
'���أ�strLabel = Ƥ�Խ����ע����"(+)"
'      intResult=Ƥ�Խ����0-���ԣ�1-����
    Dim arr���� As Variant, arr���� As Variant
    Dim i As Integer
    
    strLabel = "": intResult = 0
    
    arr���� = Split(Split(strScript, ";")(0), ",")
    arr���� = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 0: Exit Sub
        End If
    Next
End Sub

Public Function GetPriceGradeSQL(ByVal strҩƷ�۸�ȼ� As String, ByVal str���ļ۸�ȼ� As String, ByVal str��ͨ��Ŀ�۸�ȼ� As String, ByVal strTableTmpA As String, ByVal strTableTmpB As String, _
           ByVal strParNumҩƷ As String, ByVal strParNum���� As String, ByVal strParNum��ͨ��Ŀ As String) As String
'���ܣ����˼۸�ȼ����������ȡ�۸��SQL
'������strҩƷ�۸�ȼ�  '���˵�ҩƷ�۸�ȼ�
'      str���ļ۸�ȼ�  '���˵����ļ۸�ȼ�
'      str��ͨ��Ŀ�۸�ȼ�  '���˵���ͨ��Ŀ�۸�ȼ�
'     strTableTmpA   �շ���ĿĿ¼ ���as ��־,strTableTmpB  �շѼ�Ŀ�� ��As��־��
'     strParNumҩƷ  ҩƷ�۸�ȼ�SQL�������,strParNum����  ���ļ۸�ȼ�SQL�������,strParNum��ͨ��Ŀ  ��ͨ��Ŀ�۸�ȼ�SQL�������
    Dim strSQL As String
    
    If strҩƷ�۸�ȼ� = "" And str���ļ۸�ȼ� = "" And str��ͨ��Ŀ�۸�ȼ� = "" Then
        strSQL = " And " & strTableTmpB & ".�۸�ȼ� is Null "
    Else
        strSQL = " And" & vbNewLine & _
                "      ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".��� || ';') > 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNumҩƷ & "]) Or" & vbNewLine & _
                "      (Instr(';4;', ';' || " & strTableTmpA & ".��� || ';') > 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNum���� & "]) Or" & vbNewLine & _
                "      (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".��� || ';') = 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNum��ͨ��Ŀ & "]) Or" & vbNewLine & _
                "      (" & strTableTmpB & ".�۸�ȼ� Is Null And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "         From �շѼ�Ŀ" & vbNewLine & _
                "         Where " & strTableTmpA & ".Id = �շ�ϸĿid  And" & vbNewLine & _
                "               ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".��� || ';') > 0 And �۸�ȼ� = [" & strParNumҩƷ & "]) Or" & vbNewLine & _
                "               (Instr(';4;', ';' || " & strTableTmpA & ".��� || ';') > 0 And �۸�ȼ� = [" & strParNum���� & "]) Or" & vbNewLine & _
                "               (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".��� || ';') = 0 And �۸�ȼ� = [" & strParNum��ͨ��Ŀ & "]))))) "

    End If
    
    GetPriceGradeSQL = strSQL
End Function

Private Sub InitPriceLevel()
'���ܣ�����ʼ�۸�ȼ�
    Dim objTmpExpense As Object
    
    If objTmpExpense Is Nothing Then
        On Error Resume Next
        Set objTmpExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not objTmpExpense Is Nothing Then
            Call objTmpExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not objTmpExpense Is Nothing Then
        Call objTmpExpense.zlGetPriceGrade(zl9ComLib.gstrNodeNo, 0, 0, "", gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�)
    End If
End Sub

Public Sub PlugInFunc()
    '��ҳ�������ʼ��
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, 1264, -1)
        Call zlPlugInErrH(Err, "Initialize")
    End If
End Sub

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub
