Attribute VB_Name = "mdl����������"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Private Const strFolder As String = "C:\CQYB_YH"        '����Ŀ¼
Private Const strRecipe As String = "Recipe.txt"        '������ϸ
Private Const strBalance As String = "Balance.txt"      '������Ϣ
Private Const strDeal As String = "Deal.txt"            '������Ϣ
Private Const str���ﴦ����ϸ As String = "Upload.txt"  '�����������
Private mobjFileSystem As New FileSystemObject
Private mobjStream As TextStream
Public gcn���������� As New ADODB.Connection
Private mstrBusiness As String
Private mstrInput As String
Private mstrAppMsg As String
Public gstrReturn_���������� As String                            'ȫ��ʹ��
Public Const gstrSplit_Row_���������� As String = "$"
Public Const gstrSplit_Col_���������� As String = "|"
        
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Private gobjYH As Object   '���������ö���ı�����
'Private gobjYH As New clsT_CQYHYB                       '������
Private mblnInit As Boolean
Private mstrOwner As String                             '�м���û���

Private Type ComInfo_����������
    ҽԺ���� As String
    ҵ������ As String
    ���˱�� As String
    ������ˮ�� As String
    ������ˮ�� As String
    �������� As String                      '���������֤�󷵻صļ�������
    ����֢ As String
    ͳ������ As String
    �ʻ���� As Currency
    �ܷ��� As Currency                      'HIS
    �ܷ���_���� As Currency                 '���ĵķ����ܶ�
    ����ʱ�� As String
    ����ID As Long
End Type
Public gComInfo_���������� As ComInfo_����������

Enum ��������_����������
    ������Ϣ
    ������ϸ
    ������Ϣ
    ���ﴦ����ϸ
End Enum
Private rsRecipe As New ADODB.Recordset                 '�����������ﴦ��

'���½ṹ��������¼��������������ڽ���ʱ�˶�
Private Type typBalance
    curҽ������ As Double
    cur����Ա���� As Double
    cur�����ʻ� As Double
    cur�󲡻��� As Double
End Type
Private pre_Balance As typBalance

Private Function MakeFile_Recipe(ByVal rsDetail As ADODB.Recordset, Optional ByVal bln���� As Boolean = True, _
    Optional ByVal blnԤ���� As Boolean = True, Optional ByRef str������ˮ��_UP As String) As Boolean
    'str������ˮ��_Up:������¼���β����˴�����ϸ���㲿�ֵ���ˮ�ţ���","�ָ�
    Dim intDO As Integer
    Dim lng����ID As Long, lng��ҳID As Long
    Dim bln���� As Boolean, blnҩƷ As Boolean, bln�ϴ� As Boolean, blnѪҺ�׵��� As Boolean, bln���շ���Ŀ As Boolean
    Dim strҵ�� As String, str���˱�� As String, str��ˮ�� As String, str������ˮ�� As String, strͳ������ As String
    Dim str������ˮ�� As String, str�˵�������ˮ�� As String
    Dim str��Ŀ��ˮ�� As String, str��Ŀ��� As String, str�շ���� As String
    Dim strҽԺ��Ŀ���� As String, strҽԺ��Ŀ���� As String
    Dim str���� As String, str���� As String, str��� As String
    Dim str������ As String, str����ʱ�� As String, str����ҽ�� As String
    '----��Ҫ��ʼ��Ϊ��----
    Dim str��װ���� As String, str��װ��λ As String, str���� As String, str������λ As String, str���� As String, str������λ As String, str���� As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsVerify As New ADODB.Recordset
    
    '----���ڲ�ӳ���¼�����----
    Dim strFields As String, strValues As String
    
    On Error GoTo errHand
    '��ʼ���ڲ���¼��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  20      ���˱��
'    3.      string  20      ����������ˮ��
'    4.      string  18      ���㽻����ˮ��
'    5.      string  20      �˵���Ӧ����������ˮ��
'    6.      datetime        ��  ��������
'    7.      string  14      ҽ����Ŀ��ˮ��
'    8.      string  3       ��Ŀ���
'    9.      string  20      ҽԺ����
'    10.     string  50      ��Ŀ����
'    11.     number  10  4   ����
'    12.     number  8   2   ����
'    13.     number  10  4   ���
'    14.     string  50      ����
'    15.     number  8   2   ��װ����
'    16.     string  40      ��װ��λ
'    17.     string  8       ����
'    18.     string  40      ������λ
'    19.     string  8       ����
'    20.     string  14      ������λ
'    21.     string  3       �����־
'    22.     string  18      ������
'    23.     string  20      ����ҽ��
'    24.     string  20      ������
'    25.     datetime        ��  ����ʱ��
'    26.     string  3       ͳ������
    Call DebugTool("�����ڲ���¼��")
    strFields = "��ˮ��," & adLongVarChar & "," & 20 & "|���˱��," & adLongVarChar & "," & 20 & _
                "|����������ˮ��," & adLongVarChar & "," & 20 & "|���㽻����ˮ��," & adLongVarChar & "," & 20 & _
                "|�˵�����������ˮ��," & adLongVarChar & "," & 20 & "|����ʱ��," & adLongVarChar & "," & 18 & _
                "|��Ŀ��ˮ��," & adLongVarChar & "," & 15 & "|��Ŀ���," & adLongVarChar & "," & 3 & _
                "|ҽԺ��Ŀ����," & adLongVarChar & "," & 20 & "|ҽԺ��Ŀ����," & adLongVarChar & "," & 50 & _
                "|����," & adLongVarChar & "," & 18 & "|����," & adLongVarChar & "," & 18 & _
                "|���," & adLongVarChar & "," & 18 & "|����," & adLongVarChar & "," & 50 & _
                "|��װ����," & adLongVarChar & "," & 18 & "|��װ��λ," & adLongVarChar & "," & 40 & _
                "|����," & adLongVarChar & "," & 8 & "|������λ," & adLongVarChar & "," & 40 & _
                "|����," & adLongVarChar & "," & 8 & "|������λ," & adLongVarChar & "," & 14 & _
                "|����," & adLongVarChar & "," & 3 & "|������," & adLongVarChar & "," & 18 & _
                "|����ҽ��," & adLongVarChar & "," & 20 & "|������," & adLongVarChar & "," & 20 & _
                "|����ʱ��," & adLongVarChar & "," & 18 & "|ͳ������," & adLongVarChar & "," & 3
    Call Record_Init(rsRecipe, strFields)
    strFields = ""
    For intDO = 0 To rsRecipe.Fields.Count - 1
        strFields = strFields & "|" & rsRecipe.Fields(intDO).Name
    Next
    strFields = Mid(strFields, 2)
    
    '��δ�ϴ��Ĵ�����ϸ����Ϊ�����ļ�
    With rsDetail
        '��ȡ�ò��˵���ˮ�źͽ�����ˮ��
        Call DebugTool("����ֵ")
        lng����ID = !����ID
        str��ˮ�� = gComInfo_����������.������ˮ��
        str������ˮ�� = gComInfo_����������.������ˮ��
        strͳ������ = gComInfo_����������.ͳ������
        str���˱�� = gComInfo_����������.���˱��
        strҵ�� = gComInfo_����������.ҵ������
        str������ˮ��_UP = ""
        bln���� = (strҵ�� = "14")
        
        'ȡ��ҳID
        gstrSQL = "Select Nvl(סԺ����,0) AS ��ҳID From ������Ϣ Where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
        lng��ҳID = rsTemp!��ҳID
        
        '���������ļ�
        For intDO = 1 To 2
            Call DebugTool("���˷�����ϸ��������¼���󸺼�¼")
            If intDO = 1 Then
                If Not bln���� Then
                    .Filter = "���>0"
                Else
                    .Filter = "ʵ�ս��>0"
                End If
            Else
                If Not bln���� Then
                    .Filter = "���<0"
                Else
                    .Filter = "ʵ�ս��<0"
                End If
            End If
            Do While Not .EOF
                bln�ϴ� = True      '������Զ����
                If Not bln���� Then bln�ϴ� = (Nvl(!�Ƿ��ϴ�, 0) = 0)
                
                If bln�ϴ� Then
                    'ȡ������ˮ�ż��˵�������ˮ��
                    Call DebugTool("ȡ������ˮ�ż��˵�������ˮ��")
                    If bln���� And blnԤ���� Then
                        Call Get������ˮ��("", "1", "1", .AbsolutePosition, str������ˮ��, str�˵�������ˮ��, lng����ID)
                    Else
                        If bln���� Then
                            Call Get������ˮ��(!NO, !��¼����, !��¼״̬, !���, str������ˮ��, str�˵�������ˮ��, lng����ID)
                        Else
                            Call Get������ˮ��(!NO, !��¼����, !��¼״̬, !���, str������ˮ��, str�˵�������ˮ��)
                        End If
                    End If
                    str������ˮ��_UP = str������ˮ��_UP & ",'" & str������ˮ�� & "'"
                    
                    '��ֻ��ҩƷ��Ŀ���е�������Ϊ��
                    str��װ���� = "": str��װ��λ = "": str���� = "": str������λ = "": str���� = "": str������λ = "": str���� = ""
                    
                    Call DebugTool("��ȡҽ����Ŀ����������")
                    gstrSQL = "Select ���,����,���� From �շ�ϸĿ Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ��Ŀ�ı���������", CLng(!�շ�ϸĿID))
                    strҽԺ��Ŀ���� = rsTemp!����
                    strҽԺ��Ŀ���� = rsTemp!����
                    str�շ���� = rsTemp!���
                    
                    'ȡ����Ŀ��ҽ����Ϣ
                    Call DebugTool("ȡ����Ŀ��ҽ����Ϣ")
                    If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                        blnҩƷ = True
                        gstrSQL = " Select ��ˮ��,��Ŀ����,����,��װ����,��װ��λ,����,������λ,����,������λ" & _
                                  " From " & mstrOwner & ".�м��_ҩƷĿ¼ Where ��ˮ��=" & _
                                  "     (Select ��Ŀ���� From ����֧����Ŀ " & _
                                  "     Where �շ�ϸĿID=" & !�շ�ϸĿID & " And ����=" & TYPE_���������� & ")"
                    Else
                        blnҩƷ = False
                        gstrSQL = " Select ��ˮ��,��Ŀ��� ��Ŀ����" & _
                                  " From " & mstrOwner & ".�м��_������Ŀ Where ��ˮ��=" & _
                                  "     (Select ��Ŀ���� From ����֧����Ŀ " & _
                                  "     Where �շ�ϸĿID=[1] And ����=[2])"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿҽ�������Ϣ", CLng(!�շ�ϸĿID), TYPE_����������)
                    If rsItem.EOF Then
                        MsgBox "[" & strҽԺ��Ŀ���� & "]" & strҽԺ��Ŀ���� & "�е���ϸ��¼δ�ҵ���Ӧ�ı�����Ŀ�����飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    str��Ŀ��ˮ�� = Nvl(rsItem!��ˮ��)
                    str��Ŀ��� = Nvl(rsItem!��Ŀ����)
                    If blnҩƷ Then
                        str���� = Nvl(rsItem!����)
                        str��װ���� = Nvl(rsItem!��װ����)
                        str���� = Nvl(rsItem!����)
                        str������λ = Nvl(rsItem!������λ)
                        str���� = Nvl(rsItem!����)
                        str������λ = Nvl(rsItem!������λ)
                    End If
                    
                    If bln���� Then
                        str���� = Format(!ʵ�ս�� / Nvl(!����, 1), "#####0.0000;-#####0.0000; ;")
                        str���� = Format(!����, "#####0.00;-#####0.00; ;")
                        str��� = Format(!ʵ�ս��, "#####0.0000;-#####0.0000; ;")
                        
                        str����ʱ�� = Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss")
                        str����ҽ�� = Nvl(!������)
                        str������ = gComInfo_����������.������ˮ��
                    Else
                        str���� = Format(!��� / !����, "#####0.0000;-#####0.0000; ;")
                        str���� = Format(!����, "#####0.00;-#####0.00; ;")
                        str��� = Format(!���, "#####0.0000;-#####0.0000; ;")
                        
                        str����ʱ�� = Format(!����ʱ��, "yyyyMMdd HH:mm:ss")
                        str����ҽ�� = Nvl(!ҽ��)
                        str������ = !NO
                    End If
                    
                    '�˷�ʱ���俪��ʱ��ȡԭʼ������ϸ�Ŀ���ʱ��
                    Call DebugTool("�˷�ʱ���俪��ʱ��ȡԭʼ������ϸ�Ŀ���ʱ��")
                    If str�˵�������ˮ�� <> str������ˮ�� Then
                        gstrSQL = "Select to_char(��������,'yyyyMMdd hh24:mi:ss') �������� From " & mstrOwner & ".�м��_������ϸ Where ������ˮ��=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡԭʼ�Ŀ���ʱ��", str�˵�������ˮ��)
                        If Not rsTemp.EOF Then      '���Ϊ�գ���ʾסԺ�������ϣ��������ڲ���
                            str����ʱ�� = rsTemp!��������
                        End If
                    End If
                    
                    '�����סԺ�����Ǹ��շ���Ŀ��ѪҺ�׵��ף��������Ŀ���в������ݣ�ͬʱ�������ݲ��ϴ������ϴ���־����Ϊ��:bln�ϴ�=False
                    '����������е����⣺�����ϴ�ʱ������Ҫ��������Ŀ�������м��㣬��Ȼ�м��Ĵ�����ϸ����û�д��������������
                    '��˳���͸��ݴ�����������ü�¼������������������Ȼ���ٽ��м������ϴ�
                    If Not bln���� Then
                        Call ���ýӿ�_׼��_����������("100", str��Ŀ��ˮ��)
                        If Not ���ýӿ�_����������() Then Exit Function
                        
                        '20060829 �ϵ�
                        blnѪҺ�׵��� = (gstrReturn_���������� <> "")
                        bln���շ���Ŀ = (Format(str����, "#0.00") >= "1000.00" And str�շ���� <> "F")
                        If blnѪҺ�׵��� Or bln���շ���Ŀ Then
                            '�������Ŀ����˱�־�ǲ�����µ�
                            gstrSQL = "zlYB_�����Ŀ��_UPDATE(" & IIf(blnѪҺ�׵���, 2, 1) & "," & lng����ID & "," & lng��ҳID & ",'" & str������ˮ�� & "',0)"
                            gcn����������.Execute gstrSQL, , adCmdStoredProc
                        End If
                    End If
                    
                    '����������ϸ�ļ�
                    Call DebugTool("����������ϸ�ļ�")
                    strValues = str��ˮ�� & "|" & str���˱�� & "|" & str������ˮ�� & "|" & str������ˮ�� & "|" & _
                                str�˵�������ˮ�� & "|" & str����ʱ�� & "|" & str��Ŀ��ˮ�� & "|" & str��Ŀ��� & "|" & _
                                strҽԺ��Ŀ���� & "|" & strҽԺ��Ŀ���� & "|" & str���� & "|" & str���� & "|" & str��� & "|" & _
                                str���� & "|" & str��װ���� & "|" & str��װ��λ & "|" & str���� & "|" & str������λ & "|" & _
                                str���� & "|" & str������λ & "|" & IIf(bln����, "1", "0") & "|" & str������ & "|" & _
                                str����ҽ�� & "|" & gstrUserName & "|" & Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss") & "|" & _
                                strͳ������
                    Call Record_Add(rsRecipe, strFields, strValues)
                End If
                .MoveNext
            Loop
        Next
    End With
    
    If Not MakeFile_Recipe2() Then Exit Function
    If str������ˮ��_UP <> "" Then str������ˮ��_UP = Mid(str������ˮ��_UP, 2)
    MakeFile_Recipe = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Get������ˮ��(ByVal strNO As String, ByVal str���� As String, ByVal str״̬ As String, _
            ByVal str��� As String, str������ˮ�� As String, str�˵�������ˮ�� As String, _
            Optional ByVal lng����ID As Long = 0)
    Dim bln���� As Boolean
    Dim rsHandback As New ADODB.Recordset
    
    Call DebugTool("�õ�������ˮ��")
    '���صĴ�����ˮ�Ź���NO[8]+����[3]+״̬[3]+���[4] ���20λ��Ŀǰֻ�õ�18λ
    'ֻҪ�����˲���ID����˵��������
    If strNO = "" Then      '˵��������Ԥ����
        str������ˮ�� = ToVarchar(lng����ID & Mid(Format(zlDatabase.Currentdate(), "YYYYMMDDHHmmss"), 11), 15)
        str������ˮ�� = str������ˮ�� & str���
        str�˵�������ˮ�� = str������ˮ��
        Exit Sub
    End If
    
    str������ˮ�� = strNO & String(3 - Len(str����), "0") & str���� & _
                    String(3 - Len(str״̬), "0") & IIf(str״̬ = "3", "1", str״̬) & _
                    String(3 - Len(str���), "0") & str���
    If str״̬ = 1 Then
        'ȡ�ñ���ϸ�ĵ���������С���㣬�����ȡһ��������¼����ˮ����Ϊ�˵���ˮ��
        gstrSQL = " Select ����ID,��ҳID,�շ�ϸĿID,Nvl(��׼����,0) ����,NVl(ʵ�ս��,0) ���" & _
                  " From סԺ���ü�¼" & _
                  " Where NO=[1] And ��¼����=[2] And ��¼״̬=[3] And ���=[4]"
        Set rsHandback = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ϸ�����Ƿ��Ǹ�������", strNO, str����, str״̬, str���)
        If rsHandback!���� < 0 Or rsHandback!��� < 0 Then
            str�˵�������ˮ�� = GetSequence(rsHandback!����ID, rsHandback!��ҳID, rsHandback!�շ�ϸĿID)
        Else
            str�˵�������ˮ�� = str������ˮ��
        End If
    Else
        If lng����ID <> 0 Then
            '����
            '���ʣ��ӱ��ս����¼��ȡ��ժҪ������ԭʼ�Ĵ�����ˮ�ţ�
            gstrSQL = " Select ժҪ From ������ü�¼ " & _
                      " Where ����ID=[1] And ���=[2]"
            Set rsHandback = zlDatabase.OpenSQLRecord(gstrSQL, "�ӱ��ս����¼��ȡ��ժҪ", gComInfo_����������.����ID, CLng(str���))
            str�˵�������ˮ�� = rsHandback!ժҪ
        Else
            str�˵�������ˮ�� = strNO & String(3 - Len(str����), "0") & str���� & _
                            String(3 - Len(str״̬), "0") & "1" & _
                            String(3 - Len(str���), "0") & str���
        End If
    End If
End Sub

Private Function MakeFile_Recipe2() As Boolean
    '���ݼ�¼�����������ļ������ʽ�͸�����ˮ����ȡ���д�����ϸ���������ļ���һ��
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(������ϸ) Then Exit Function
    
    With rsRecipe
        lngCols = .Fields.Count - 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_Recipe2 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_RecipeCalculated(ByVal str��ˮ�� As String, Optional ByVal str������ˮ�� As String) As Boolean
    '���ӿڴ�����Ĵ�����ϸ���м������ȡ���������ξ����������ϸ����������Ϊ�����ļ�
    Dim lngCol As Long, lngCols As Long
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'ȡ����ID����ҳID
    gstrSQL = " Select A.����ID,A.סԺ���� AS ��ҳID From ������Ϣ A,�����ʻ� B" & _
              " Where A.����ID=B.����ID ANd B.��ˮ��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID����ҳID", str��ˮ��)
    lng����ID = rsTemp!����ID
    lng��ҳID = rsTemp!��ҳID
    
    If Not CreateExchangeFile(������ϸ) Then Exit Function
    '20060829 �ϵ�
    gstrSQL = " SELECT ��ˮ��,���˱��,������ˮ��,ҽ����Ŀ����,ҽ����Ŀ��ˮ��,���㽻����ˮ��,�˵�������ˮ��,ҽ�����, " & _
              "     ҽ�ƻ�������,��Ŀ����,��Ŀ����,�����־,���շ��������,������־,����,����, " & _
              "     ���,���Ը����,������,��Ŀ���,����޼�,����,��װ����,��װ��λ,����, " & _
              "     ������λ,����,������λ,����ҽ��,to_char(��������,'yyyyMMdd hh24:mi:ss') ��������,������,to_char(����ʱ��,'yyyyMMdd hh24:mi:ss') ����ʱ��,ͳ������,��ע  " & _
              " FROM �м��_������ϸ" & _
              " Where ��ˮ��='" & str��ˮ�� & "' And ������ˮ�� Not In " & _
              "     (Select ������ˮ�� From �����Ŀ�� Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & " And ��˱�־=0)" & _
              IIf(str������ˮ�� = "", "", " And ������ˮ�� in (" & str������ˮ�� & ")")
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����������
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_RecipeCalculated = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_Deal(ByVal str��ˮ�� As String) As Boolean
    '���м������ȡָ����ˮ�ŵĴ�����Ϣ��ֻ������һ������������Ϊ�����ļ�
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(������Ϣ) Then Exit Function
    gstrSQL = "SELECT �������,������¼���,��ˮ��,���˱��,ҽ�����,��Ա���,��������,ʵ������, " & _
            "     ���ܹ���Ա����,�����סԺ��־,�����������,����ԭ��,ҽ�ƻ����ȼ�,ת��ǰҽ�ƻ�������,���ֱ���,�ز��������, " & _
            "     ���������Ƿ���סԺ,��������סԺ����,��������סԺ��ߵȼ�,���������ۼ�סԺ����,�����ʻ����,�𸶱�׼, " & _
            "     ��������,���������������ۼ�,ͳ��֧���ۼ�,����Ա��������ۼ�,�ز�����ҽ�����ۼ�,��ʷδ�������Ը�, " & _
            "     תԺǰ�ѽ���ͳ��,תԺǰ����Ա��������,to_char(��ʼʱ��,'yyyyMMdd hh24:mi:ss') ��ʼʱ��,to_char(����ʱ��,'yyyyMMdd hh24:mi:ss') ����ʱ��,ͳ������  " & _
            " FROM �м��_ҽ�ƴ�����Ϣ" & _
            " Where ��ˮ��='" & str��ˮ�� & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����������
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_Deal = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_Balance(ByVal str��ˮ�� As String) As Boolean
    '���м������ȡָ����ˮ�ţ����ξ�������ν����¼��������Ϊ�����ļ�
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim blnEmpty As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(������Ϣ) Then Exit Function
    
    blnEmpty = True
    gstrSQL = "SELECT ��ˮ��,���㽻����ˮ��,�˵�������ˮ��,������¼���,���˱��,����,��Ա���, " & _
             "  ���ܹ���Ա,ҽ�ƻ�������,ҽ�ƻ����ȼ�,ҽ�����,���ֱ���,��������,��������, " & _
             "  ���ⲡ֢��־,�ز��������,��������,����סԺ����,�ز��Ը�����,����סԺ��������, " & _
             "  ҽ�Ʒ��ܶ�,�Է��ܶ�,�����ʻ�֧���ܶ�,�����ֽ�֧���ܶ�,�����ؼ��Ը��ܶ�, " & _
             "  ����ҩ�Ը��ܶ�,����Ա�������Ը�,����Ա��������,���Ը����ֹ���Ա����, " & _
             "  ��ʷ���Ը�����Ա����,����ʵ��֧������,�𸶱�׼�Ը����,�����¹���Ա����, " & _
             "  �����¹���Ա����,��ʷ���߹���Ա����,������ͨ���﹫��Ա����,���Ϸ�Χҽ����, " & _
             "  ��һ�ν��,��һ���Ը�����,�ڶ��ν��,�ڶ����Ը�����,�����ν��,�������Ը�����, " & _
             "  �ֶ��Ը����,���λ���ͳ����,�ֶ��Ը�����Ա����,����󲡽��,��֧�����, " & _
             "  תԺ��������ҽ����,תԺ�����������,תԺ�������빫��Ա,תԺ���������, " & _
             "  תԺ���������Ը�,��Ʊ��,������,to_char(����ʱ��,'yyyyMMdd hh24:mi:ss') ����ʱ��,ͳ������  " & _
             " FROM �м��_������Ϣ" & _
             " Where trim(��ˮ��)='" & Trim(str��ˮ��) & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����������
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            blnEmpty = False
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    
    If blnEmpty Then mobjStream.WriteLine ""
    mobjStream.Close
    
    MakeFile_Balance = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Deal(Optional ByVal blnSave As Boolean = False) As Boolean
    '�����ӿڷ��صĴ����ļ��������浽�м�⣨Ԥ���㷵�صĽ��80%��׼ȷ����˽��鲻���棩
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, strBuffer As String
    Dim lngRow As Long
    Dim arrCol
    
    Const int��ʼ����  As Integer = 30
    Const int�������� As Integer = 31
    On Error GoTo errHand
    
'    ҽ�ƴ�����Ϣ(�������,������¼���,��ˮ��,���˱��,ҽ�����,��Ա���,��������,ʵ������,���ܹ���Ա����,�����סԺ��־,
'        �����������,����ԭ��,ҽ�ƻ����ȼ�,ת��ǰҽ�ƻ�������,���ֱ���,�ز��������,���������Ƿ���סԺ,��������סԺ����,
'        ��������סԺ��ߵȼ�,���������ۼ�סԺ����,�����ʻ����,�𸶱�׼,��������,���������������ۼ�,ͳ��֧���ۼ�,
'        ����Ա��������ۼ�,�ز�����ҽ�����ۼ�,��ʷδ�������Ը�,תԺǰ�ѽ���ͳ��,תԺǰ����Ա��������,��ʼʱ��,����ʱ��,ͳ������)
    Call DebugTool("����������Ϣ�ļ�")
    strData = "ZL_�м��_ҽ�ƴ�����Ϣ_Insert("
    If Not OpenExchangeFile(������Ϣ) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        lngRow = mobjStream.Line
        strBuffer = mobjStream.ReadLine
        strDeal = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            Select Case lngCol
            Case int��ʼ����, int��������
                '�������ڸ�ʽ��ͬ����Ҫת��
                strDate = arrCol(lngCol)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                strDeal = strDeal & strDate
            Case Else
                strDeal = strDeal & ",'" & arrCol(lngCol) & "'"
            End Select
        Next
        strDeal = strData & Mid(strDeal, 2) & IIf(lngRow = 1, ",1", "") & ")"
        If blnSave Then gcn����������.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_Deal = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Recipe(Optional ByVal blnSave As Boolean = False) As Boolean
    '�����ӿڷ��صĴ�����ϸ�ļ��������浽�м�⣨Ԥ���㷵�صĽ��80%��׼ȷ����˽��鲻���棩
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strRecipe As String, strBuffer As String
    Dim arrCol
    
    Const int�������� As Integer = 29
    Const int�������� As Integer = 31
    On Error GoTo errHand
    
    '������ϸ(��ˮ��,���˱��,������ˮ��,ҽ����Ŀ����,ҽ����Ŀ��ˮ��,���㽻����ˮ��,�˵�������ˮ��,
    '       ҽ�����,ҽ�ƻ�������,��Ŀ����,��Ŀ����,�����־,���շ��������,������־,����,
    '       ����,���,���Ը����,������,��Ŀ���,����޼�,����,��װ����,��װ��λ,����,������λ,
    '       ����,������λ,����ҽ��,��������,������,����ʱ��,ͳ������,��ע)
    Call DebugTool("����������ϸ�ļ�")
    strData = "ZL_�м��_������ϸ_Insert("
    If Not OpenExchangeFile(������ϸ) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        strBuffer = mobjStream.ReadLine
        strRecipe = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            Select Case lngCol
            Case int��������, int��������
                '�������ڸ�ʽ��ͬ����Ҫת��
                strDate = arrCol(lngCol)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                strRecipe = strRecipe & strDate
            Case Else
                strRecipe = strRecipe & ",'" & arrCol(lngCol) & "'"
            End Select
        Next
        strRecipe = strData & Mid(strRecipe, 2) & ")"
        If blnSave Then gcn����������.Execute strRecipe, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_Recipe = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Balance(strReturn As String, Optional ByVal blnSave As Boolean = False) As Boolean
    '�����ӿڷ��صĽ������ļ��������浽�м�⣨Ԥ���㷵�صĽ��80%��׼ȷ����˽��鲻���棩
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strBalance As String, strBuffer As String
    Dim arrCol
    Dim curҽ������ As Currency, cur����Ա���� As Currency, cur�����ʻ� As Currency, cur�󲡻��� As Currency
    
    Const int�����ܶ� As Integer = 20
    Const int����ͳ�� As Integer = 44
    Const int����Ա1 As Integer = 28
    Const int����Ա2 As Integer = 29
    Const int����Ա3 As Integer = 33
    Const int����Ա4 As Integer = 34
    Const int����Ա5 As Integer = 35
    Const int����Ա6 As Integer = 45
    Const int����Ա7 As Integer = 50
    Const int�����ʻ� As Integer = 22
    Const int��ͳ�� As Integer = 47
    Const int����ʱ�� As Integer = 55
    On Error GoTo errHand
    
'    ������Ϣ(��ˮ��,���㽻����ˮ��,�˵�������ˮ��,������¼���,���˱��,����,��Ա���,���ܹ���Ա,ҽ�ƻ�������,
'        ҽ�ƻ����ȼ�,ҽ�����,���ֱ���,��������,��������,���ⲡ֢��־,�ز��������,��������,����סԺ����,
'        �ز��Ը�����,����סԺ��������,ҽ�Ʒ��ܶ�,�Է��ܶ�,�����ʻ�֧���ܶ�,�����ֽ�֧���ܶ�,�����ؼ��Ը��ܶ�,
'        ����ҩ�Ը��ܶ�,����Ա�������Ը�[26],����Ա��������,���Ը����ֹ���Ա����[28],��ʷ���Ը�����Ա����[29],����ʵ��֧������,
'        �𸶱�׼�Ը����,�����¹���Ա����,�����¹���Ա����[33],��ʷ���߹���Ա����[34],������ͨ���﹫��Ա����[35],
'        ���Ϸ�Χҽ����[36],��һ�ν��[37],��һ���Ը�����,�ڶ��ν��[39],�ڶ����Ը�����,�����ν��[41],�������Ը�����,
'        �ֶ��Ը����,���λ���ͳ����[44],�ֶ��Ը�����Ա����[45],����󲡽��,��֧�����[47],תԺ��������ҽ����,
'        תԺ�����������,תԺ�������빫��Ա,תԺ���������,תԺ���������Ը�,��Ʊ��,������,����ʱ��,ͳ������)
    gComInfo_����������.�ܷ���_���� = 0
    Call DebugTool("����������Ϣ�ļ�")
    strData = "ZL_�м��_������Ϣ_Insert("
    If Not OpenExchangeFile(������Ϣ) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        strBuffer = mobjStream.ReadLine
        strBalance = ""
        If Trim(strBuffer) <> "" Then
            arrCol = Split(strBuffer, vbTab)
            lngCols = UBound(arrCol)
            For lngCol = 0 To lngCols
                Select Case lngCol
                Case int����ʱ��
                    '�������ڸ�ʽ��ͬ����Ҫת��
                    strDate = arrCol(lngCol)
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                    strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    strBalance = strBalance & strDate
                Case Else
                    strBalance = strBalance & ",'" & arrCol(lngCol) & "'"
                End Select
            Next
            strBalance = strData & Mid(strBalance, 2) & ")"
            If blnSave Then gcn����������.Execute strBalance, , adCmdStoredProc
        
            '��ȡÿ�ʼ�¼��ҽ��ͳ�����Ա�������������ʻ�֧���ܶ�
            gComInfo_����������.�ܷ���_���� = gComInfo_����������.�ܷ���_���� + Val(arrCol(int�����ܶ�))
            curҽ������ = curҽ������ + Val(arrCol(int����ͳ��))
            cur����Ա���� = cur����Ա���� + Val(arrCol(int����Ա1)) + Val(arrCol(int����Ա2)) + Val(arrCol(int����Ա3)) + _
                    Val(arrCol(int����Ա4)) + Val(arrCol(int����Ա5)) + Val(arrCol(int����Ա6)) + Val(arrCol(int����Ա7))
            cur�����ʻ� = cur�����ʻ� + Val(arrCol(int�����ʻ�))
            cur�󲡻��� = cur�󲡻��� + Val(arrCol(int��ͳ��))
        End If
    Loop
    mobjStream.Close
    
    If curҽ������ <> 0 Then strReturn = strReturn & "|" & "ҽ������;" & curҽ������ & ";0"
    If cur����Ա���� <> 0 Then strReturn = strReturn & "|" & "����Ա��������;" & cur����Ա���� & ";0"
    If cur�����ʻ� <> 0 Then strReturn = strReturn & "|" & "�����ʻ�;" & cur�����ʻ� & ";0"
    If cur�󲡻��� <> 0 Then strReturn = strReturn & "|" & "�󲡻���;" & cur�󲡻��� & ";0"
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    If strReturn = "" Then strReturn = "�����ʻ�;0;0"
    AnalyFile_Balance = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenExchangeFile(ByVal int���� As ��������_����������) As Boolean
    '���ļ�
    Dim strFile As String
    On Error GoTo errHand
    
    strFile = GetFileName(int����)
    If Not mobjFileSystem.FileExists(strFile) Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile(strFile, ForReading, False, TristateMixed)
    
    OpenExchangeFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CreateExchangeFile(ByVal int���� As ��������_����������) As Boolean
    On Error GoTo errHand
    
    Set mobjStream = mobjFileSystem.CreateTextFile(GetFileName(int����))
    
    CreateExchangeFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetFileName(ByVal int���� As ��������_����������) As String
    Select Case int����
    Case ��������_����������.������ϸ
        GetFileName = strRecipe
    Case ��������_����������.������Ϣ
        GetFileName = strDeal
    Case ��������_����������.������Ϣ
        GetFileName = strBalance
    Case ��������_����������.���ﴦ����ϸ
        GetFileName = str���ﴦ����ϸ
    End Select
    GetFileName = strFolder & "\" & GetFileName
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim intRecords As Integer
    '������:����
    '��������:2000-11-02
    'Ҳʹ���ڱ���
    Set RecTarget = New ADODB.Recordset
    
    With RecTarget
        If .State = 1 Then .Close
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, adLongVarChar, 100, adFldIsNullable     '0:��ʾ����
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        Do While Not SourceRec.EOF
            If Nvl(SourceRec!�Ƿ��ϴ�, 0) = 0 Then
                .AddNew
                For intFields = 0 To SourceRec.Fields.Count - 1
                    .Fields(intFields) = SourceRec.Fields(intFields).Value
                Next
                .Update
            End If
            If Nvl(SourceRec!�Ƿ��ϴ�, 0) = 0 Then
                intRecords = intRecords + 1
                If intRecords = 15 Then
                    SourceRec.MoveNext
                    Exit Do
                End If
            End If
            SourceRec.MoveNext
        Loop
    End With
    
    Set CopyNewRec = RecTarget
End Function

Public Function ��ݱ�ʶ_����������(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim str��ˮ�� As String, StrInput As String, strIdentify As String
    Dim blnTrans As Boolean
    Dim strReturn As String
    Dim arrReturn
    On Error GoTo errHand
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    strIdentify = frmIdentify����������.GetPatient(bytType, lng����ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0) Then Exit Function
    
    '��������
    gcn����������.BeginTrans
    blnTrans = True
    
    '���������ҵ������þ���Ǽǽӿ�
    If bytType = 0 Then
        '1.      string  20      ��ᱣ�Ϻ�
        '2.      string  20      ����/סԺ��
        '3.      string  3       ҽ����𣬼������
        '4.      string  30      ����
        '5.      string  20      ҽ��
        '6.      datetime        ��  ��Ժ����
        '7.      string  20      ��Ժ����
        '8.      string  20      ������
        '9.      string  50      ����֢
        '10.     string          ������ҽ�ƴ�����Ϣ�ļ������·�����ļ���
        gComInfo_����������.����ʱ�� = Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00"
        str��ˮ�� = ToVarchar(lng����ID & Format(zlDatabase.Currentdate(), "yyMMddHHmmss"), 18)
        StrInput = gComInfo_����������.���˱�� & gstrSplit_Col_���������� & str��ˮ�� & gstrSplit_Col_���������� & _
                 gComInfo_����������.ҵ������ & gstrSplit_Col_���������� & "����" & gstrSplit_Col_���������� & _
                 ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & gComInfo_����������.����ʱ�� & gstrSplit_Col_���������� & _
                 gComInfo_����������.�������� & gstrSplit_Col_���������� & ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & _
                 ToVarchar(gComInfo_����������.����֢, 50) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
        Call ���ýӿ�_׼��_����������("08", StrInput)
        If Not ���ýӿ�_����������() Then
            gcn����������.RollbackTrans
            Exit Function
        End If
        strReturn = gstrReturn_����������
        If Not AnalyFile_Deal(True) Then
            gcn����������.RollbackTrans
            Exit Function
        End If
        
        '�õ�������ˮ�źͽ�����ˮ�ţ����ϲ�����ȷ��ɲű����µ���ˮ�ţ�
        arrReturn = Split(strReturn, gstrSplit_Col_����������)
        gComInfo_����������.������ˮ�� = arrReturn(1)
        gComInfo_����������.������ˮ�� = arrReturn(0)
        
        '���½��㽻����ˮ�ż�����/סԺ��ˮ��
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'��ˮ��','''" & gComInfo_����������.������ˮ�� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������ˮ��")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'������ˮ��','''" & gComInfo_����������.������ˮ�� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������㽻����ˮ��")
    End If
    
    '���±����ʻ������Ϣ��ͳ�����š�ҵ�����ͣ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'ͳ������','''" & gComInfo_����������.ͳ������ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ͳ������")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'ҵ������','''" & gComInfo_����������.ҵ������ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'����֢','''" & gComInfo_����������.����֢ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���沢��֢")
    
    gcn����������.CommitTrans
    
    '���ز�����Ϣ��
    ��ݱ�ʶ_���������� = strIdentify
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn����������.RollbackTrans
End Function

Public Function ҽ����ʼ��_����������(Optional ByVal blnTest As Boolean = False) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    '����Ƿ���ڽ���Ŀ¼���������򴴽�
    If Not mobjFileSystem.FolderExists(strFolder) Then
        mobjFileSystem.CreateFolder (strFolder)
    End If
    
    If mblnInit = False Then
        If Not blnTest Then '����ǲ��ԣ���˵���Ǳ��ղ������ô�����
            '��������ҽ��������������
            gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_����������)
            
            Do Until rsTemp.EOF
                Select Case rsTemp("������")
                    Case "ҽ���û���"
                        strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ��������"
                        strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ���û�����"
                        strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                End Select
                rsTemp.MoveNext
            Loop
            
            mstrOwner = strUser
            If OraDataOpen(gcn����������, strServer, strUser, strPass, False) = False Then
                MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Set gobjYH = CreateObject("YinHai.ChongQing.MedicareDefray")
        '��������Ƿ���
        If gobjYH Is Nothing Then
            MsgBox "ҽ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            '��������ҽ�������� 204-04-07
            Exit Function
        End If
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_����������)
        gComInfo_����������.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        If Not blnTest Then mblnInit = True
        
        'У������ʱ�䣬�������ӿ�ֻȡ����ʱ�䣨�����������ݾ�ȷ���룬��ʽΪ��yyyyMMdd HH:mm:ss��
        Call ���ýӿ�_׼��_����������("01")
        On Error Resume Next
        gstrReturn_���������� = Mid(gstrReturn_����������, 10)
        If Err = 0 Then Time = gstrReturn_����������
    End If
    
    ҽ����ʼ��_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ������_����������() As Boolean
    ҽ������_���������� = frmSet����������.��������
End Function

Public Function ҽ����ֹ_����������() As Boolean
    On Error Resume Next
    
    Set gobjYH = Nothing
    gcn����������.Close
    Set gcn���������� = Nothing
    
    mblnInit = False
    ҽ����ֹ_���������� = True
End Function

Public Sub ���ýӿ�_׼��_����������(ByVal strBusiness As String, Optional ByVal StrInput As String = "", _
    Optional ByVal strOutput As String = "", Optional ByVal strAppMsg As String = "")
    mstrAppMsg = strAppMsg
    mstrBusiness = strBusiness
    mstrInput = StrInput
    gstrReturn_���������� = strOutput
End Sub

Public Function ���ýӿ�_����������() As Boolean
    '���״���    ��������    �Ƿ���Ҫ������
    '1   ��ȡԶ��ϵͳʱ��       '��
    '2   ��ȡҩƷĿ¼           '
    '3   ��ȡ������ĿĿ¼       '
    '4   ��ȡ����Ŀ¼           '
    '5   ��ȡҽ��������Ϣ       '
    '6   ��ȡ���������Ϣ       '
    '7   ��ȡ���˻�����Ϣҽ�ƴ�����Ϣ    '��
    '8   ����Ǽ�               '��
    '9   ������Ϣ�޸�           '��
    '10  ������ϸ����           '��
    '11  ��ȡ���շ���Ŀ������Ϣ '��
    '12  ������ϸ��Ϣ�ϴ�       '��
    '13  ģ����ý���           '
    '14  ���ý���               '��
    '15  ����Ǽ�����           '��
    '16  �˶Է��ý�����Ϣ       '��
    '17  �˶Դ�����ϸ��Ϣ       '��
    '18  ��ȡҩƷĿ¼��ʷ�����Ϣ    '
    '19  ��ȡ������ĿĿ¼��ʷ�����Ϣ    '
    '20  ��ȡ����Ŀ¼��ʷ�����Ϣ    '
    '22  ��ȡ������Ϣ�������ϴν���ʧ�ܣ����ĳɹ���HISʧ�ܵ������
    '23  ��ȡ������Ϣ
    '33  �����Ϣд��
    On Error GoTo errHand
    Dim lngResult As Long
    
    Call DebugTool(String(20, "-"))
    Call DebugTool("���״��룺" & mstrBusiness)
    Call DebugTool("��Σ�" & mstrInput)
    lngResult = gobjYH.passivebusiness(mstrBusiness, mstrInput, gstrReturn_����������, mstrAppMsg)
    If lngResult < 0 Then               '������Ϣ
        MsgBox "������ʾ����������[" & mstrBusiness & "]�������[" & lngResult & "]" & mstrAppMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf lngResult > 0 Then           '������Ӧ����ʾ��Ϣ
        MsgBox "������ʾ��" & mstrAppMsg, vbInformation, gstrSysName
    End If
    
    ���ýӿ�_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �����������_����������(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strFileName As String, StrInput As String, strReturn As String
    On Error GoTo errHand
    '�õ����ν�����ܷ���
    Call DebugTool("�õ����ν�����ܷ���")
    gComInfo_����������.�ܷ��� = 0
    With rs��ϸ
        Do While Not .EOF
            gComInfo_����������.�ܷ��� = gComInfo_����������.�ܷ��� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '���ϴ���ϸ���ȵ��ô�����ϸ���㣬�ٵ���Ԥ����
    '----������ϸ����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string          ����Ĵ�����ϸ��Ϣ�ļ������·�����ļ���
'    OutputString
'    ���    ��������    ����    ����    ˵��
'    1.      string          �����Ĵ�����ϸ��Ϣ�ļ������·�����ļ���
    Call DebugTool("׼������������ϸ�ļ����Ա�ӿڼ���")
    If Not MakeFile_Recipe(rs��ϸ, True, True) Then Exit Function
    strFileName = GetFileName(������ϸ)
    Call DebugTool("���ô�����ϸ����ӿ�")
    Call ���ýӿ�_׼��_����������("10", strFileName & gstrSplit_Col_���������� & strFileName)
    If Not ���ýӿ�_���������� Then Exit Function
    If Not AnalyFile_Recipe Then Exit Function
    
    '----ģ�����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  18      ���㽻����ˮ��
'    3.      string          ҽ�ƴ�����Ϣ�ļ������·�����ļ���
'    4.      string          ������ϸ��Ϣ�ļ������·�����ļ���
'    5.      string          �ôξ������η��ý������ļ������·�����ļ���
'    6.      string          ���ý������ļ������·�����ļ���
    '�Ȳ���������漰�����ļ����ٵ��ýӿڣ�ע�⣺������ϸ�ļ�ֱ��ʹ�ô�����ϸ����󣬽ӿڷ��ص��ļ����ɣ��������²�����
    Call DebugTool("����������Ϣ�ļ�")
    If Not MakeFile_Deal(gComInfo_����������.������ˮ��) Then Exit Function
    If Not MakeFile_Balance(gComInfo_����������.������ˮ��) Then Exit Function
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������ϸ) & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call DebugTool("����13�ӿڣ�����ģ�����")
    Call ���ýӿ�_׼��_����������("13", StrInput)
    If Not ���ýӿ�_���������� Then Exit Function
    Call DebugTool("�����������ļ�")
    If Not AnalyFile_Balance(strReturn) Then Exit Function
    str���㷽ʽ = strReturn
    Call AnalyBalance(str���㷽ʽ)
    
    �����������_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_����������(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, Optional ByRef strAdvance As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    Dim LNGDO As Long, lngLoop As Long              'ѭ�����������Կ��ƴ�����ϸ�ϴ�
    Dim intCounter As Integer                       '������
    Dim lng����ID As Long
    Dim blnTrans As Boolean
    Dim StrInput As String, strReturn As String, strBillNO As String, str������ˮ�� As String, strFileName As String
    Dim curҽ������ As Currency, cur�󲡻��� As Currency, cur����Ա�������� As Currency, cur�ֽ� As Currency, curMoney As Currency
    Dim intBalance As Integer, intBalances As Integer, str���㷽ʽ As String, arrBalance
    Dim str����ʱ�� As String, str��Ժԭ�� As String
    Dim objStream As TextStream, objFileSys As New FileSystemObject
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim blnOld As Boolean, blnRevise As Boolean '�Ƿ���Ҫ��дУ���ֶΣ��������Ƿ���ҪУ��
    On Error GoTo errHand
    
    '��Ԥ���㲻һ������Ҫ�ڵ��ô�����ϸ����󣬽��ŵ��ô�����ϸ�ϴ�������ٵ��ý���
    gcn����������.BeginTrans
    blnTrans = True
    'ȡ���������з�����ϸ
    gstrSQL = "Select ID From ������ü�¼ Where ����ID=" & lng����ID & " And Nvl(��¼״̬,0)<>0 Order by ���"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���������з�����ϸ")
    '����ѭ����������Ϊÿ���ļ��ϴ��ļ�¼��ֻ��15��
    lngLoop = (rs��ϸ.RecordCount \ 15) + IIf(rs��ϸ.RecordCount Mod 15 = 0, 0, 1)
    
    If Not AnalyFile_Recipe(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
'    '�������δͨ�������ĸ��շ���Ŀ
'    If Not CheckItem Then
'        If Not frm�ȴ���Ӧ_����������.ShowME() Then
'            gcn����������.RollbackTrans
'            Exit Function
'        End If
'    End If
'
'    '�ٴβ�����ϸ�ļ�
'    rsRecipe.MoveFirst
'    If Not MakeFile_Recipe2() Then
'        gcn����������.RollbackTrans
'        Exit Function
'    End If
'
'    '�ٴε��ô�����ϸ���㣬�Ի�ȡ�µĴ�����ϸ�ļ����Ա�����������
'    strFileName = GetFileName(������ϸ)
'    Call ���ýӿ�_׼��_����������("10", strFileName & gstrSplit_Col_���������� & strFileName)
'    If Not ���ýӿ�_���������� Then
'        gcn����������.RollbackTrans
'        Exit Function
'    End If
'    If Not AnalyFile_Recipe(True) Then
'        gcn����������.RollbackTrans
'        Exit Function
'    End If
    
    '���ݴ�����ϸ�ļ����²��˷��ü�¼��ժҪ���������洦����ˮ�ţ�
    With rsRecipe
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs��ϸ!ID & ",NULL,NULL,NULL,NULL,1,'" & rsRecipe!����������ˮ�� & "')"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
            rs��ϸ.MoveNext
        Loop
    End With
    
    '----������ϸ�ϴ�----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string          ����Ĵ�����ϸ��Ϣ�ļ������·�����ļ���
    Set objStream = objFileSys.OpenTextFile(GetFileName(������ϸ))
    For LNGDO = 1 To lngLoop
        intCounter = 0
        '��ԭʼ�ļ������ϴ�������ϸ�ļ������20����ϸ��
        If Not CreateExchangeFile(���ﴦ����ϸ) Then
            gcn����������.RollbackTrans
            Exit Function
        End If
        
        Do While Not objStream.AtEndOfStream
            If intCounter = 15 Then Exit Do
            mobjStream.WriteLine objStream.ReadLine
            intCounter = intCounter + 1
        Loop
        
        mobjStream.Close
        '�ϴ�����
        Call ���ýӿ�_׼��_����������("12", GetFileName(���ﴦ����ϸ))
        If Not ���ýӿ�_���������� Then
            gcn����������.RollbackTrans
            objStream.Close
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "������ϸ�ϴ�ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    objStream.Close
    
    '�õ���ʼ��Ʊ��
    gstrSQL = "Select ����ID,ʵ��Ʊ�� From ������ü�¼ Where ����ID=[1] And ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ʊ��", lng����ID)
    strBillNO = Nvl(rsTemp!ʵ��Ʊ��)
    lng����ID = rsTemp!����ID
    
    '----��ʽ����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  18      ���㽻����ˮ��
'    3.      string  3       ��������  0-��������;1-��;����;2-ת��ͥ��������;3-����ʧ����
'    4.      string  20      ���ֱ���
'    5.      string  20      ��Ʊ��
'    6.      string          ������ϸ��Ϣ�ļ������·�����ļ���
'    7.      string          �ôξ������η��ý������ļ������·�����ļ���
'    8.      string          ����ҽ�ƴ�����Ϣ�ļ������·�����ļ���
'    9.      string          �������ý������ļ������·�����ļ���
'    OutputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ���㽻����ˮ��
    '�Ȳ���������漰�����ļ����ٵ��ýӿڣ�ע�⣺������ϸ�ļ�ֱ��ʹ�ô�����ϸ����󣬽ӿڷ��ص��ļ����ɣ��������²�����
    If Not MakeFile_Deal(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & _
             "0" & gstrSplit_Col_���������� & gComInfo_����������.�������� & gstrSplit_Col_���������� & _
             strBillNO & gstrSplit_Col_���������� & GetFileName(������ϸ) & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("14", StrInput)
    
    If Not ���ýӿ�_���������� Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    str������ˮ�� = gstrReturn_����������
    If Not AnalyFile_Deal(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '�ֽ����֧����Ϣ
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str���㷽ʽ = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str���㷽ʽ
        Case "�����ʻ�"
            cur�����ʻ� = curMoney
        Case "ҽ������"
            curҽ������ = curMoney
        Case "�󲡻���"
            cur�󲡻��� = curMoney
        Case "����Ա��������"
            cur����Ա�������� = curMoney
        End Select
    Next
    cur�ֽ� = gComInfo_����������.�ܷ��� - cur�����ʻ� - curҽ������ - cur�󲡻��� - cur����Ա��������
    
    '�Խ��������к˶�
    If Not (cur�����ʻ� = pre_Balance.cur�����ʻ� And curҽ������ = pre_Balance.curҽ������ And _
        cur�󲡻��� = pre_Balance.cur�󲡻��� And cur����Ա�������� = pre_Balance.cur����Ա����) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    '���±����ʻ����µĽ�����ˮ�ţ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'������ˮ��','''" & str������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㽻����ˮ��")
    
    '���汾�ν������
    '����ͳ����=����Ա��������;ͳ�ﱨ�����=ͳ�����;���Ը����=�󲡻���
    '��ע=ҵ������|������ˮ��|������ˮ��|����ʱ��|��������|����֢
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���������� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_����������.�ܷ��� & "," & cur�ֽ� & "," & 0 & "," & cur����Ա�������� & "," & curҽ������ & "," & cur�󲡻��� & "," & _
        0 & "," & cur�����ʻ� & ",null,null,null,'" & gComInfo_����������.ҵ������ & "|" & gComInfo_����������.������ˮ�� & "|" & gComInfo_����������.������ˮ�� & "|" & gComInfo_����������.����ʱ�� & "|" & gComInfo_����������.�������� & "|" & gComInfo_����������.����֢ & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    '����Ǽ��޸ģ���ʶΪ����
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  3       ҽ�����
'    3.      string  30      ����
'    4.      string  20      ҽ��
'    5.      datetime        ��  ��Ժ����
'    6.      string  20      ��Ժ����
'    7.      string  3       ����״̬
'    8.      datetime        ��  ��Ժ����
'    9.      string  20      ȷ�Ｒ������
'    10.     string  3       ��Ժԭ��
'    11.     string  20      ������
'    12.     string  50      ����֢
    str����ʱ�� = Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00"
    str��Ժԭ�� = 1
    
    'ȡҵ�������Ϣ
    gstrSQL = " Select A.��ˮ��,A.ҵ������,B.���� ��������,A.����֢ From �����ʻ� A,���ղ��� B " & _
              " Where A.����ID=[1] And A.����=[2] And A.����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ", lng����ID, TYPE_����������)
    
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.ҵ������ & gstrSplit_Col_���������� & _
            "����" & gstrSplit_Col_���������� & ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & _
            str����ʱ�� & gstrSplit_Col_���������� & gComInfo_����������.�������� & gstrSplit_Col_���������� & _
            "0" & gstrSplit_Col_���������� & str����ʱ�� & gstrSplit_Col_���������� & _
            gComInfo_����������.�������� & gstrSplit_Col_���������� & str��Ժԭ�� & gstrSplit_Col_���������� & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & ToVarchar(gComInfo_����������.����֢, 50)
    Call ���ýӿ�_׼��_����������("09", StrInput)
    If Not ���ýӿ�_����������() Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
   '����Ԥ����Ľ��80%��������ʽ����Ľ����һ�£���Ϊ������Ԥ����ӿڲ���ȥȡ�������µĴ�����Ϣ�������Ҫ������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur�󲡻��� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�󲡻���|" & cur�󲡻���
    If cur����Ա�������� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա��������|" & cur����Ա��������
    If str���㷽ʽ <> "" And blnRevise Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        If blnOld Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    gcn����������.CommitTrans
    
    blnTrans = False
    �������_���������� = True
'
'    '��ӡƱ��
'    Call ���ýӿ�_׼��_����������("21", gComInfo_����������.������ˮ��)
'    Call ���ýӿ�_����������
    
    Exit Function
errHand:
    If blnTrans Then gcn����������.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_����������(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim blnTrans As Boolean, bln�ϴ� As Boolean
    Dim lng����ID As Long
    Dim StrInput As String, strReturn As String, strBillNO As String, str������ˮ�� As String, strFileName As String
    Dim curҽ������ As Currency, cur�󲡻��� As Currency, cur����Ա�������� As Currency, cur�ֽ� As Currency, curMoney As Currency
    Dim intBalance As Integer, intBalances As Integer, str���㷽ʽ As String, arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim LNGDO As Long, lngLoop As Long              'ѭ�����������Կ��ƴ�����ϸ�ϴ�
    Dim intCounter As Integer                       '������
    Dim objStream As TextStream, objFileSys As New FileSystemObject
    On Error GoTo errHand
    
    '��Ҫ��ȡ�ϴξ������ػ�����Ϣ��������ˮ�š�ҵ�����͡����ֱ���ȣ�
    '���ս����¼.��ע=ҵ������|������ˮ��|������ˮ��|����ʱ��|��������|����֢
    gstrSQL = " Select B.ҽ���� ���˱��,A.֧��˳���,A.��ע,B.������ˮ�� From ���ս����¼ A,�����ʻ� B " & _
              " Where A.����=1 And A.��¼ID=[1] And A.����ID=B.����ID And B.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴν���ʱ�ľ�����ˮ��", lng����ID, TYPE_����������)
    gComInfo_����������.������ˮ�� = Split(rsTemp!��ע, "|")(1)
    gComInfo_����������.ҵ������ = Split(rsTemp!��ע, "|")(0)
    gComInfo_����������.����ʱ�� = Split(rsTemp!��ע, "|")(3)
    gComInfo_����������.������ˮ�� = rsTemp!������ˮ��  'ֻ�д�����ȡ��ǰ�Ľ�����ˮ��
    gComInfo_����������.�������� = Split(rsTemp!��ע, "|")(4)
    gComInfo_����������.����֢ = Split(rsTemp!��ע, "|")(5)
    gComInfo_����������.���˱�� = rsTemp!���˱��
    gComInfo_����������.����ID = lng����ID
    
    '��Ԥ���㲻һ������Ҫ�ڵ��ô�����ϸ����󣬽��ŵ��ô�����ϸ�ϴ�������ٵ��ý���
    gcn����������.BeginTrans
    blnTrans = True
    gComInfo_����������.�ܷ��� = 0
    
    'ȡ������¼�Ľ���ID�����ݺ�
    gstrSQL = "select distinct A.����ID,A.NO from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    strBillNO = rsTemp!NO
    
    '����ϸ��¼��
    gstrSQL = " Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,C.��Ŀ���� ҽ����Ŀ���� ,A.ʵ�ս��,A.����*Nvl(A.����,1) ����,A.ʵ�ս��/(A.����*Nvl(A.����,1)) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
              " From ������ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & TYPE_���������� & ") C " & _
              " Where A.��¼����=1 And A.��¼״̬=2 And A.NO=[1]" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
              " Order by A.NO,A.����ID"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ��¼", strBillNO)
    With rs��ϸ
        Do While Not .EOF
            gComInfo_����������.�ܷ��� = gComInfo_����������.�ܷ��� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        lngLoop = (rs��ϸ.RecordCount \ 20) + IIf(rs��ϸ.RecordCount Mod 20 = 0, 0, 1)
    End With
    
    '����������ϸ�����õĽ����ļ�����ΪNO����Ԥ����ʱ��û�У�����޷�ʹ��Ԥ����ӿڷ��صĴ�����ϸ�ļ���
    '----������ϸ����----
    strBillNO = "Z9000999"      '�˷�ʱ�޷�Ʊ��
    If Not MakeFile_Recipe(rs��ϸ, True, False) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    strFileName = GetFileName(������ϸ)
    Call ���ýӿ�_׼��_����������("10", strFileName & gstrSplit_Col_���������� & strFileName)
    If Not ���ýӿ�_���������� Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Recipe(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '----������ϸ�ϴ�----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string          ����Ĵ�����ϸ��Ϣ�ļ������·�����ļ���
    Set objStream = objFileSys.OpenTextFile(GetFileName(������ϸ))
    For LNGDO = 1 To lngLoop
        intCounter = 0
        '��ԭʼ�ļ������ϴ�������ϸ�ļ������20����ϸ��
        If Not CreateExchangeFile(���ﴦ����ϸ) Then
            gcn����������.RollbackTrans
            Exit Function
        End If
        
        Do While Not objStream.AtEndOfStream
            If intCounter = 20 Then Exit Do
            mobjStream.WriteLine objStream.ReadLine
            intCounter = intCounter + 1
        Loop
        
        mobjStream.Close
        '�ϴ�����
        Call ���ýӿ�_׼��_����������("12", GetFileName(���ﴦ����ϸ))
        If Not ���ýӿ�_���������� Then
            gcn����������.RollbackTrans
            objStream.Close
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "������ϸ�ϴ�ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    objStream.Close
    
    With rs��ϸ
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
        'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        .MoveNext
    End With
    
    gcn����������.CommitTrans
    gcn����������.BeginTrans
    '----��ʽ����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  18      ���㽻����ˮ��
'    3.      string  3       ��������  0-��������;1-��;����;2-ת��ͥ��������;3-����ʧ����
'    4.      string  20      ���ֱ���
'    5.      string  20      ��Ʊ��
'    6.      string          ������ϸ��Ϣ�ļ������·�����ļ���
'    7.      string          �ôξ������η��ý������ļ������·�����ļ���
'    8.      string          ����ҽ�ƴ�����Ϣ�ļ������·�����ļ���
'    9.      string          �������ý������ļ������·�����ļ���
'    OutputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ���㽻����ˮ��
    '�Ȳ���������漰�����ļ����ٵ��ýӿڣ�ע�⣺������ϸ�ļ�ֱ��ʹ�ô�����ϸ����󣬽ӿڷ��ص��ļ����ɣ��������²�����
    If Not MakeFile_RecipeCalculated(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Deal(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & _
             "0" & gstrSplit_Col_���������� & gComInfo_����������.�������� & gstrSplit_Col_���������� & _
             strBillNO & gstrSplit_Col_���������� & GetFileName(������ϸ) & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("14", StrInput)
    If Not ���ýӿ�_���������� Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    str������ˮ�� = gstrReturn_����������
    If Not AnalyFile_Deal(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '�ֽ����֧����Ϣ
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str���㷽ʽ = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str���㷽ʽ
        Case "�����ʻ�"
            cur�����ʻ� = curMoney
        Case "ҽ������"
            curҽ������ = curMoney
        Case "�󲡻���"
            cur�󲡻��� = curMoney
        Case "����Ա��������"
            cur����Ա�������� = curMoney
        End Select
    Next
    cur�ֽ� = gComInfo_����������.�ܷ��� - cur�����ʻ� - curҽ������ - cur�󲡻��� - cur����Ա��������
    
    '���±����ʻ����µĽ�����ˮ�ţ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'������ˮ��','''" & str������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㽻����ˮ��")
    
    '���汾�ν������
    '����ͳ����=����Ա��������;ͳ�ﱨ�����=ͳ�����;���Ը����=�󲡻���
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���������� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_����������.�ܷ��� & "," & cur�ֽ� & "," & 0 & "," & cur����Ա�������� & "," & curҽ������ & "," & cur�󲡻��� & "," & _
        0 & "," & cur�����ʻ� & ",null,null,null,'" & gComInfo_����������.ҵ������ & "|" & gComInfo_����������.������ˮ�� & "|" & gComInfo_����������.������ˮ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    gcn����������.CommitTrans
    ����������_���������� = True
    
    '��ӡƱ��
'    Call ���ýӿ�_׼��_����������("21", gComInfo_����������.������ˮ��)
'    Call ���ýӿ�_����������
    Exit Function
errHand:
    If blnTrans Then gcn����������.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��Ժ�Ǽ�_����������(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str��ˮ�� As String, StrInput As String, strReturn As String
    Dim arrReturn
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '1.      string  20      ��ᱣ�Ϻ�
    '2.      string  20      ����/סԺ��
    '3.      string  3       ҽ����𣬼������
    '4.      string  30      ����
    '5.      string  20      ҽ��
    '6.      datetime        ��  ��Ժ����
    '7.      string  20      ��Ժ����
    '8.      string  20      ������
    '9.      string  50      ����֢
    '10.     string          ������ҽ�ƴ�����Ϣ�ļ������·�����ļ���
    gcn����������.BeginTrans
    blnTrans = True
    
    gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.���� ����,A.����ҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
    
    gComInfo_����������.����ʱ�� = Format(rsTemp!��Ժ����, "yyyyMMdd") & " 00:00:00"
    str��ˮ�� = ToVarchar(lng����ID & "_" & lng��ҳID, 18)
    StrInput = gComInfo_����������.���˱�� & gstrSplit_Col_���������� & str��ˮ�� & gstrSplit_Col_���������� & _
             gComInfo_����������.ҵ������ & gstrSplit_Col_���������� & rsTemp!���� & gstrSplit_Col_���������� & _
             Nvl(rsTemp!ҽ��, "����") & gstrSplit_Col_���������� & gComInfo_����������.����ʱ�� & gstrSplit_Col_���������� & _
             gComInfo_����������.�������� & gstrSplit_Col_���������� & ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & _
             ToVarchar(gComInfo_����������.����֢, 50) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("08", StrInput)
    If Not ���ýӿ�_����������() Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    strReturn = gstrReturn_����������
    If Not AnalyFile_Deal(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '�õ�������ˮ�źͽ�����ˮ�ţ����ϲ�����ȷ��ɲű����µ���ˮ�ţ�
    arrReturn = Split(strReturn, gstrSplit_Col_����������)
    gComInfo_����������.������ˮ�� = arrReturn(1)
    gComInfo_����������.������ˮ�� = arrReturn(0)
    
    '���½��㽻����ˮ�ż�����/סԺ��ˮ��
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'��ˮ��','''" & gComInfo_����������.������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������ˮ��")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'������ˮ��','''" & gComInfo_����������.������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㽻����ˮ��")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    gcn����������.CommitTrans
    
    ��Ժ�Ǽ�_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn����������.RollbackTrans
End Function

Public Function ��Ժ�Ǽǳ���_����������(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "��ҽ�����˴���δ����ã��������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    '����Ƿ��Ѵ��ڷ��ü�¼,������ֻ��������Ժ�Ǽ�
    gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���ڷ��ü�¼", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        MsgBox "�ò����Ѿ����ڷ��ü�¼,ֻ�ܰ����Ժ������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡԭ������ˮ��
    gstrSQL = "Select ��ˮ�� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ˮ��", TYPE_����������, lng����ID)
    gComInfo_����������.������ˮ�� = rsTemp!��ˮ��
    
    '���þ���Ǽ����Ͻӿ�
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
    Call ���ýӿ�_׼��_����������("15", gComInfo_����������.������ˮ��)
    If Not ���ýӿ�_����������() Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_����������(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim blnҽ����Ժ As Boolean, bln���� As Boolean
    Dim str����ʱ�� As String, str���� As String, strҽ�� As String, str��Ժԭ�� As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
'    akc195  1   ��Ժԭ��    ����
'    akc195  2   ��Ժԭ��    תԺ
'    akc195  3   ��Ժԭ��    ����
'    akc195  4   ��Ժԭ��    ��ת
'    akc195  9   ��Ժԭ��    ����
    
    blnҽ����Ժ = False
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '�жϸò����Ƿ�������û�н�����Ĳ��˷���Ϊ�㣬˵����Ҫ���þ���Ǽǳ���
        bln���� = False
        gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(����ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�õ��þ���Ǽǳ���", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            bln���� = True
        End If
        
        blnҽ����Ժ = True
        If bln���� Then
            'ȡ��Ժ�����Ϣ
            gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.���� ����,A.��Ժ��ʽ,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
                      " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
            str����ʱ�� = Format(rsTemp!��Ժ����, "yyyyMMdd") & " 00:00:00"
            str���� = Nvl(rsTemp!����)
            strҽ�� = IIf(IsNull(rsTemp!ҽ��), "����", rsTemp!ҽ��)
            str��Ժԭ�� = IIf(IsNull(rsTemp!��Ժ��ʽ), "", rsTemp!��Ժ��ʽ)
            Select Case str��Ժԭ��
            Case "����", "����"
                str��Ժԭ�� = 1
            Case "����"
                str��Ժԭ�� = 3
            Case "תԺ"
                str��Ժԭ�� = 2
            Case Else
                str��Ժԭ�� = 9
            End Select
            
            'ȡҵ�������Ϣ
            gstrSQL = " Select A.��ˮ��,A.ҵ������,B.���� ��������,A.����֢ From �����ʻ� A,���ղ��� B " & _
                      " Where A.����ID=[1] And A.����=[2] And A.����ID=B.ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ", lng����ID, TYPE_����������)
            
            StrInput = rsTemp!��ˮ�� & gstrSplit_Col_���������� & rsTemp!ҵ������ & gstrSplit_Col_���������� & _
                    str���� & gstrSplit_Col_���������� & strҽ�� & gstrSplit_Col_���������� & _
                    str����ʱ�� & gstrSplit_Col_���������� & rsTemp!�������� & gstrSplit_Col_���������� & _
                    "0" & gstrSplit_Col_���������� & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_���������� & _
                    rsTemp!�������� & gstrSplit_Col_���������� & str��Ժԭ�� & gstrSplit_Col_���������� & _
                    ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & ToVarchar(Nvl(rsTemp!����֢), 50)
            Call ���ýӿ�_׼��_����������("09", StrInput)
            If Not ���ýӿ�_����������() Then Exit Function
        Else
            '��ȡԭ������ˮ��
            gstrSQL = "Select ��ˮ�� From �����ʻ� Where ����=[1] And ����ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ˮ��", TYPE_����������, lng����ID)
            gComInfo_����������.������ˮ�� = rsTemp!��ˮ��
            
            '���þ���Ǽ����Ͻӿ�
        '    InputString
        '    ���    ��������    ����    ����    ˵��
        '    1.      string  18      ����/סԺ��ˮ��
            Call ���ýӿ�_׼��_����������("15", gComInfo_����������.������ˮ��)
            If Not ���ýӿ�_����������() Then Exit Function
        End If
    End If
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    MsgBox IIf(blnҽ����Ժ, "ҽ����Ժ", "HIS��Ժ") & "����ɹ���", vbInformation, gstrSysName
    ��Ժ�Ǽ�_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����������(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim StrInput As String
    Dim str����ʱ�� As String, str���� As String, strҽ�� As String, str��Ժԭ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        'ȡ��Ժ�����Ϣ
        gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.���� ����,A.��Ժ��ʽ,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
                  " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
        str����ʱ�� = Format(rsTemp!��Ժ����, "yyyyMMdd") & " 00:00:00"
        str���� = rsTemp!����
        strҽ�� = Nvl(rsTemp!ҽ��, "����")
        str��Ժԭ�� = IIf(IsNull(rsTemp!��Ժ��ʽ), "", rsTemp!��Ժ��ʽ)
        Select Case str��Ժԭ��
        Case "����", "����"
            str��Ժԭ�� = 1
        Case "����"
            str��Ժԭ�� = 3
        Case "תԺ"
            str��Ժԭ�� = 2
        Case Else
            str��Ժԭ�� = 9
        End Select
        
        'ȡҵ�������Ϣ
        gstrSQL = " Select A.��ˮ��,A.ҵ������,B.���� ��������,A.����֢ From �����ʻ� A,���ղ��� B " & _
                  " Where A.����ID=[1] And A.����=[2] And A.����ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ", lng����ID, TYPE_����������)
        
        StrInput = rsTemp!��ˮ�� & gstrSplit_Col_���������� & rsTemp!ҵ������ & gstrSplit_Col_���������� & _
                str���� & gstrSplit_Col_���������� & strҽ�� & gstrSplit_Col_���������� & _
                str����ʱ�� & gstrSplit_Col_���������� & rsTemp!�������� & gstrSplit_Col_���������� & _
                "1" & gstrSplit_Col_���������� & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_���������� & _
                rsTemp!�������� & gstrSplit_Col_���������� & str��Ժԭ�� & gstrSplit_Col_���������� & _
                ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & ToVarchar(Nvl(rsTemp!����֢), 50)
        Call ���ýӿ�_׼��_����������("09", StrInput)
        If Not ���ýӿ�_����������() Then Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_����������(strSelfNo As String) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim strReturn As String
    Const int�ʻ���� As Integer = 13
    On Error GoTo errHandle
    
    'ֱ�ӵ��������֤��ȡ�������
    Call ���ýӿ�_׼��_����������("07", strSelfNo)
    If Not ���ýӿ�_���������� Then Exit Function
    strReturn = gstrReturn_����������
    �������_���������� = Val(Split(strReturn, gstrSplit_Col_����������)(int�ʻ����))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����������(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim strFile_Recipe As String, StrInput As String, strReturn As String
    Dim str������ˮ��_UP As String                     '��¼�����ϴ���ϸ����ˮ��
    Dim str������ˮ�� As String, str�����˵���ˮ�� As String
    Dim blnTrans As Boolean, bln�ϴ� As Boolean     '�Ƿ���������,�Ƿ����δ�ϴ��ļ�¼
    Dim bln���� As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim gcn�ϴ� As New ADODB.Connection
    Dim rsRecipe As New ADODB.Recordset
    Dim intDO As Integer
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHand
    '��δ�ϴ�����ϸ���м��㣬���ϴ����д�����ϸ���ٵ���Ԥ����ӿ�
    Call DebugTool("����סԺ�������")
    'ȡ�ò��˵����ҵ����Ϣ
    gstrSQL = " Select A.ҽ����,A.��ˮ��,A.ҵ������,A.ͳ������,A.������ˮ��,A.����֢,B.���� �������� From �����ʻ� A,���ղ��� B" & _
              " Where A.����ID=[1] And A.����=[2] And A.����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵����ҵ����Ϣ", lng����ID, TYPE_����������)
    With gComInfo_����������
        .���˱�� = rsTemp!ҽ����
        .����֢ = Nvl(rsTemp!����֢)
        .������ˮ�� = rsTemp!��ˮ��
        .������ˮ�� = rsTemp!������ˮ��
        .�������� = rsTemp!��������
        .ҵ������ = rsTemp!ҵ������
        .ͳ������ = rsTemp!ͳ������
        .�ܷ��� = 0
    End With
    
    '�´�һ�����������ϴ�������ϸ�������ظ��ϴ�
    Set gcn�ϴ� = GetNewConnection
    blnTrans = True
    gcn����������.BeginTrans
    
    With rsExse
        Call DebugTool("����Ƿ���룬�һ��ܷ����ܶ�")
        Do While Not .EOF
            If Nvl(!�Ƿ��ϴ�, 0) = 0 And Not bln�ϴ� Then
                bln�ϴ� = True
                gcn�ϴ�.BeginTrans
            End If
            If Nvl(!ҽ����Ŀ����) = "" Then
                If bln�ϴ� Then gcn�ϴ�.RollbackTrans
                gcn����������.RollbackTrans
                MsgBox "����δ�������Ŀ���������ϴ���", vbInformation, gstrSysName
                Exit Function
            End If
            gComInfo_����������.�ܷ��� = gComInfo_����������.�ܷ��� + Nvl(!���, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        If bln�ϴ� Then
            Call DebugTool("�ϴ���ϸ")
            strFile_Recipe = GetFileName(������ϸ)
        
            For intDO = 1 To 2
                If intDO = 1 Then
                    .Filter = "���>0"
                Else
                    .Filter = "���<=0"
                End If
            
                '��δ�ϴ�����ϸ���м���
                Do While Not rsExse.EOF
                    Call DebugTool("����������ϸ��¼���������ϴ�")
                    If Not blnTrans Then gcn����������.BeginTrans: blnTrans = True
                    If Not bln�ϴ� Then gcn�ϴ�.BeginTrans: bln�ϴ� = True
                    Set rsRecipe = CopyNewRec(rsExse)
                    
                    If rsRecipe.RecordCount <> 0 Then
                        Call DebugTool("����������ϸ�ļ�")
                        rsRecipe.Filter = 0
                        rsRecipe.MoveFirst
                        If Not MakeFile_Recipe(rsRecipe, False, False, str������ˮ��_UP) Then
                            .Filter = 0
                            rsRecipe.Filter = 0
                            gcn����������.RollbackTrans
                            gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        
                        rsRecipe.Filter = 0
                        Call ���ýӿ�_׼��_����������("10", strFile_Recipe & gstrSplit_Col_���������� & strFile_Recipe)
                        Call DebugTool("���ô�����ϸ����")
                        If Not ���ýӿ�_���������� Then
                            .Filter = 0
                            gcn����������.RollbackTrans
                            gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        If Not AnalyFile_Recipe(True) Then
                            .Filter = 0
                            gcn����������.RollbackTrans
                            gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        
                        '----������ϸ�ϴ���������δ�ϴ����ֵĴ�����ϸ��----
                        '    InputString
                        '    ���    ��������    ����    ����    ˵��
                        '    1.      string          ����Ĵ�����ϸ��Ϣ�ļ������·�����ļ���
                        Call DebugTool("���ô�����ϸ�ϴ�")
                        '����������ϸ�ļ�
                        Call MakeFile_RecipeCalculated(gComInfo_����������.������ˮ��, str������ˮ��_UP)
                        
                        Call ���ýӿ�_׼��_����������("12", strFile_Recipe)
                        If Not ���ýӿ�_���������� Then
                            .Filter = 0
                            gcn����������.RollbackTrans
                            gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        
                        '�����ϴ��Ĵ������ϴ���־
                        Call DebugTool("���ϴ����")
                        With rsRecipe
                            If .RecordCount <> 0 Then .MoveFirst
                            Do While Not .EOF
                                If Nvl(!�Ƿ��ϴ�, 0) = 0 Then
                                    '20060829 �ϵ�
                                    bln���� = True
                                    
                                    Call Get������ˮ��(!NO, !��¼����, !��¼״̬, !���, str������ˮ��, str�����˵���ˮ��)
                                    
                                    '������Ǵ�������Ŀ������˱�־�Ѹ��µģ������ϴ���־
                                    gstrSQL = "Select ��˱�־ From �����Ŀ�� Where ������ˮ��='" & str������ˮ�� & "'"
                                    Call OpenRecordset_OtherBase(rsTemp, "�ж�", gstrSQL, gcn����������)
                                    If rsTemp.RecordCount <> 0 Then
                                        bln���� = (rsTemp!��˱�־ <> 0)
                                    End If
                                    
                                    '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                                    If bln���� Then
                                        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                                        gcn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                                    End If
                                End If
                                .MoveNext
                            Loop
                        End With
                    End If
                    
                    '��֤������ϸ���㱣��ɹ�
                    gcn�ϴ�.CommitTrans
                    gcn����������.CommitTrans
                    bln�ϴ� = False
                    blnTrans = False
                Loop
            Next
            
            .Filter = 0
        End If
    End With
    
    If blnTrans = False Then gcn����������.BeginTrans: blnTrans = True
    Call DebugTool("��ȡ���շ���Ŀ������Ϣ")
    Call TestVerifyItem
    
    '�����д�����ϸ������������ȡ������������Ϊ�����ļ���׼������Ԥ����ӿ�
    '----ģ�����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  18      ���㽻����ˮ��
'    3.      string          ҽ�ƴ�����Ϣ�ļ������·�����ļ���
'    4.      string          ������ϸ��Ϣ�ļ������·�����ļ���
'    5.      string          �ôξ������η��ý������ļ������·�����ļ���
'    6.      string          ���ý������ļ������·�����ļ���
    '�Ȳ���������漰�����ļ����ٵ��ýӿڣ�ע�⣺������ϸ�ļ�ֱ��ʹ�ô�����ϸ����󣬽ӿڷ��ص��ļ����ɣ��������²�����
    If Not MakeFile_Deal(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_RecipeCalculated(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    Call DebugTool("����סԺ�������ӿ�")
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������ϸ) & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("13", StrInput)
    If Not ���ýӿ�_���������� Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    gcn����������.CommitTrans
    סԺ�������_���������� = strReturn
    Call AnalyBalance(strReturn)
    If סԺ�������_���������� = "" Then סԺ�������_���������� = "�����ʻ�;0;0"
    
    '����ܶ�ȣ�����ʾ
    If Format(gComInfo_����������.�ܷ���, "#0.00") <> Format(gComInfo_����������.�ܷ���_����, "#0.00") Then
        MsgBox "���ֱ���δ�����ܷ�����ҽ�����Ĳ�һ�£�" & vbCrLf & _
               "ҽԺ��" & Format(gComInfo_����������.�ܷ���, "#0.00") & Space(10) & "ҽ�����ģ�" & Format(gComInfo_����������.�ܷ���_����, "#0.00"), vbInformation, gstrSysName
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn����������.RollbackTrans
    If bln�ϴ� Then gcn�ϴ�.RollbackTrans
End Function

Public Function סԺ����_����������(lng����ID As Long, ByVal lng����ID As Long, Optional ByRef strAdvance As String) As Boolean
    Dim strBillNO As String, StrInput As String, strReturn As String, str���㷽ʽ As String, str������ˮ�� As String
    Dim cur�����ʻ� As Currency, curҽ������ As Currency, cur�󲡻��� As Currency, cur����Ա�������� As Currency, cur�ֽ� As Currency, curMoney As Currency
    Dim cur�����ʻ�_OLD As Currency, curҽ������_OLD As Currency, cur�󲡻���_OLD As Currency, cur����Ա��������_OLD As Currency, cur�ֽ�_OLD As Currency
    Dim intBalance As Integer, intBalances As Integer, arrBalance
    Dim blnTrans As Boolean, bln������ As Boolean, lng����ID As Long
    Dim intState As Integer
    Dim lng��ҳID As Long
    Dim blnOld As Boolean, blnRevise As Boolean '�Ƿ���Ҫ��дУ���ֶΣ��������Ƿ���ҪУ��
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
        '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo errHand
    
    '��ȡ�������ͣ�������ڣ�˵���ǵ����ֽ���
    gstrSQL = " Select Nvl(���,0) AS ��� From ���ղ��� Where ID=" & _
              "     (Select ����ID From �����ʻ� Where ����=[1] And ����ID=[2])" & _
              " And ����=" & TYPE_����������
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ֵ�����", TYPE_����������, lng����ID)
    If rsTemp.RecordCount <> 0 Then
        bln������ = (rsTemp!��� = 4)
    End If
    
    '�õ���ʼ��Ʊ��
    gstrSQL = "Select ����ID,ʵ��Ʊ�� From ���˽��ʼ�¼ Where ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ʊ��", lng����ID)
    strBillNO = Nvl(rsTemp!ʵ��Ʊ��)
    
    '�õ���ҳID
    gstrSQL = "Select Nvl(סԺ����,0) ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵���ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    '���������Ҫ��������Ŀȴδ���������������
    '20060829 �ϵ�
    gstrSQL = "Select Count(*) From �����Ŀ�� Where ��˱�־=0 And ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
    Call OpenRecordset_OtherBase(rsTemp, "���������Ҫ��������Ŀȴδ���������������", gstrSQL, gcn����������)
    If rsTemp.Fields(0).Value > 0 Then
        MsgBox "����" & rsTemp.Fields(0).Value & "����������Ŀ����������㣡", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������  0-��������;1-��;����;2-ת��ͥ��������;3-����ʧ����;5-�����ֽ���
    intState = 1
    gstrSQL = "Select Nvl(��ǰ״̬,0) ״̬ From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�ǰ״̬", lng����ID, TYPE_����������)
    '����ǳ�Ժ���㣬����Ϊ���������־
    If rsTemp!״̬ = 0 Then intState = 0
    '������ҽ���ڣ������ֲ��˲�������;����ĸ����˲������нỹ�ǳ�Ժ���㣬����5
    If bln������ Then intState = 5
    
    If intState = 1 Then        '��;��������ʾ�Ƿ�ת��Ժ��ͥ��������
        If MsgBox("�ò����Ƿ����ת��ͥ�������㣿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            intState = 2
        End If
    End If
    
    '�����ϴν���������ĳɹ�����HISδ�ɹ�����ˣ�����ý���22��ȡ������Ϣ����������ܶ�Ϊ�㣬˵���ϴν���ɹ������ΰ��������̽��㼴��
    '----------��ȡ�ϴν�������----------
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("22", StrInput)
    If Not ���ýӿ�_���������� Then Exit Function
    If Not AnalyFile_Balance(strReturn) Then Exit Function
    
    '�ֽ����֧����Ϣ
    If gComInfo_����������.�ܷ���_���� <> 0 Then
        '��������ܶΪ�㣬��˵���ϴ��ѽ��㣬�µĽ�����ˮ���Է��ص�Ϊ׼���н���
        gComInfo_����������.������ˮ�� = gstrReturn_����������
        arrBalance = Split(strReturn, "|")
        intBalances = UBound(arrBalance)
        For intBalance = 0 To intBalances
            str���㷽ʽ = Split(arrBalance(intBalance), ";")(0)
            curMoney = Split(arrBalance(intBalance), ";")(1)
            Select Case str���㷽ʽ
            Case "�����ʻ�"
                cur�����ʻ�_OLD = curMoney
            Case "ҽ������"
                curҽ������_OLD = curMoney
            Case "�󲡻���"
                cur�󲡻���_OLD = curMoney
            Case "����Ա��������"
                cur����Ա��������_OLD = curMoney
            End Select
        Next
        If Not AnalyFile_Balance(strReturn, True) Then Exit Function
    End If
    
    gcn����������.BeginTrans
    blnTrans = True
    '----��ʽ����----
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  18      ���㽻����ˮ��
'    3.      string  3       ��������  0-��������;1-��;����;2-ת��ͥ��������;3-����ʧ����
'    4.      string  20      ���ֱ���
'    5.      string  20      ��Ʊ��
'    6.      string          ������ϸ��Ϣ�ļ������·�����ļ���
'    7.      string          �ôξ������η��ý������ļ������·�����ļ���
'    8.      string          ����ҽ�ƴ�����Ϣ�ļ������·�����ļ���
'    9.      string          �������ý������ļ������·�����ļ���
'    OutputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ���㽻����ˮ��
    '�Ȳ���������漰�����ļ����ٵ��ýӿ�
    '������ϸ��Ԥ����ʱ�Ѿ������ˣ������ٴβ���
    If Not MakeFile_Deal(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_����������.������ˮ��) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '----------��ȡ���ν�������----------
    StrInput = gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & gComInfo_����������.������ˮ�� & gstrSplit_Col_���������� & _
             intState & gstrSplit_Col_���������� & gComInfo_����������.�������� & gstrSplit_Col_���������� & _
             strBillNO & gstrSplit_Col_���������� & GetFileName(������ϸ) & gstrSplit_Col_���������� & _
             GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ) & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("14", StrInput)
    
    '����ǵ����֣��ڽ��㽻��ʹ������״̬����Ҫ��״̬��Ϊ��;������Ժ����
    If intState = 5 Then
        intState = 1
        gstrSQL = "Select Nvl(��ǰ״̬,0) ״̬ From �����ʻ� Where ����ID=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�ǰ״̬", lng����ID, TYPE_����������)
        '����ǳ�Ժ���㣬����Ϊ���������־
        If rsTemp!״̬ = 0 Then intState = 0
    End If
    
    If Not ���ýӿ�_���������� Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    str������ˮ�� = gstrReturn_����������
    If Not AnalyFile_Deal(True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    strReturn = ""
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    '�ֽ����֧����Ϣ
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str���㷽ʽ = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str���㷽ʽ
        Case "�����ʻ�"
            cur�����ʻ� = curMoney
        Case "ҽ������"
            curҽ������ = curMoney
        Case "�󲡻���"
            cur�󲡻��� = curMoney
        Case "����Ա��������"
            cur����Ա�������� = curMoney
        End Select
    Next
    
    '��Ӧ���ۼӣ����������Զ������ϴεĽ����¼���ж�������ʹ����ͬ�Ľ�����ˮ��
'    '�ۼ����εĽ������ݣ�Ϊ���εĽ�����
'    cur�����ʻ� = cur�����ʻ� + cur�����ʻ�_OLD
'    curҽ������ = curҽ������ + curҽ������_OLD
'    cur�󲡻��� = cur�󲡻��� + cur�󲡻���_OLD
'    cur����Ա�������� = cur����Ա�������� + cur����Ա��������_OLD
'    gComInfo_����������.�ܷ��� = gComInfo_����������.�ܷ��� + gComInfo_����������.�ܷ���_����
'    cur�ֽ� = gComInfo_����������.�ܷ��� - cur�����ʻ� - curҽ������ - cur�󲡻��� - cur����Ա��������
    
    '�Ƚ������������ʽ�������Ƿ�һ��
    If Not (cur�����ʻ� = pre_Balance.cur�����ʻ� And curҽ������ = pre_Balance.curҽ������ And _
        cur�󲡻��� = pre_Balance.cur�󲡻��� And cur����Ա�������� = pre_Balance.cur����Ա����) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    '���±����ʻ����µĽ�����ˮ�ţ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'������ˮ��','''" & str������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㽻����ˮ��")
    
    '���汾�ν������
    '����ͳ����=����Ա��������;ͳ�ﱨ�����=ͳ�����;���Ը����=�󲡻���
    '��ע=ҵ������|������ˮ��|������ˮ��|����ʱ��|��������|����֢
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���������� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_����������.�ܷ��� & "," & cur�ֽ� & "," & 0 & "," & cur����Ա�������� & "," & curҽ������ & "," & cur�󲡻��� & "," & _
        0 & "," & cur�����ʻ� & ",null,null,null,'" & gComInfo_����������.ҵ������ & "|" & gComInfo_����������.������ˮ�� & "|" & gComInfo_����������.������ˮ�� & "|" & gComInfo_����������.����ʱ�� & "|" & gComInfo_����������.�������� & "|" & gComInfo_����������.����֢ & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")

    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʼ�¼�����ϴ���־")
    
    '����Ԥ����Ľ��80%��������ʽ����Ľ����һ�£���Ϊ������Ԥ����ӿڲ���ȥȡ�������µĴ�����Ϣ�������Ҫ������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur�󲡻��� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�󲡻���|" & cur�󲡻���
    If cur����Ա�������� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա��������|" & cur����Ա��������
    If str���㷽ʽ <> "" And blnRevise Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        If blnOld Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
        Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    'ȡ��Ժ�����Ϣ
    Dim str����ʱ�� As String, str���� As String, strҽ�� As String, str��Ժԭ�� As String
    gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.���� ����,A.��Ժ��ʽ,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
    str����ʱ�� = Format(rsTemp!��Ժ����, "yyyyMMdd") & " 00:00:00"
    str���� = Nvl(rsTemp!����)
    strҽ�� = IIf(IsNull(rsTemp!ҽ��), "����", rsTemp!ҽ��)
    str��Ժԭ�� = IIf(IsNull(rsTemp!��Ժ��ʽ), "", rsTemp!��Ժ��ʽ)
    Select Case str��Ժԭ��
    Case "����", "����"
        str��Ժԭ�� = 1
    Case "����"
        str��Ժԭ�� = 3
    Case "תԺ"
        str��Ժԭ�� = 2
    Case Else
        str��Ժԭ�� = 9
    End Select
    
    'ȡҵ�������Ϣ
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
'    2.      string  3       ҽ�����
'    3.      string  30      ����
'    4.      string  20      ҽ��
'    5.      datetime        ��  ��Ժ����
'    6.      string  20      ��Ժ����
'    7.      string  3       ����״̬
'    8.      datetime        ��  ��Ժ����
'    9.      string  20      ȷ�Ｒ������
'    10.     string  3       ��Ժԭ��
'    11.     string  20      ������
'    12.     string  50      ����֢
    gstrSQL = " Select A.��ˮ��,A.ҵ������,B.���� ��������,A.����֢ From �����ʻ� A,���ղ��� B " & _
              " Where A.����ID=[1] And A.����=[2] And A.����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ", lng����ID, TYPE_����������)
    
    StrInput = rsTemp!��ˮ�� & gstrSplit_Col_���������� & rsTemp!ҵ������ & gstrSplit_Col_���������� & _
            str���� & gstrSplit_Col_���������� & Nvl(strҽ��, "����") & gstrSplit_Col_���������� & _
            str����ʱ�� & gstrSplit_Col_���������� & rsTemp!�������� & gstrSplit_Col_���������� & _
            "1" & gstrSplit_Col_���������� & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_���������� & _
            rsTemp!�������� & gstrSplit_Col_���������� & str��Ժԭ�� & gstrSplit_Col_���������� & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & ToVarchar(Nvl(rsTemp!����֢), 50)
    Call ���ýӿ�_׼��_����������("09", StrInput)
    If Not ���ýӿ�_����������() Then
        gcn����������.RollbackTrans
        Exit Function
    End If
    
    gcn����������.CommitTrans
    סԺ����_���������� = True
    
    'ͬʱ�����Ժ�Ǽ�
    If intState = 0 Then Call ��Ժ�Ǽ�_����������(lng����ID, lng��ҳID)
    Exit Function
errHand:
    If blnTrans Then gcn����������.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����������(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    MsgBox "ҽ����֧�ֽ������ϣ���ֱ�����ϼ��ʵ��ݺ��ٽ��ʣ�", vbInformation, gstrSysName
    סԺ�������_���������� = False
End Function

Private Function Get���ղ���_����������(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������=[1] and A.����=[2] and A.���� is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", str������, TYPE_����������)
    
    If rsTemp.EOF = False Then
        Get���ղ���_���������� = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Function �۸��ж�_����������(ByVal dblҽԺ As Double, ByVal dblҽ�� As Double, ByVal str�޼۷�ʽ As String, _
                              ByVal bln�ؼ� As Boolean, ByVal dbl�ؼ� As Double) As Boolean
'���ܣ��ж�ҽԺ�ļ۸��Ƿ񳬹�ҽ���涨�ĵ���
    Dim strҽԺ��� As String
    
    On Error GoTo errHandle
    
    If InStr(str�޼۷�ʽ, "����") > 0 Then
        strҽԺ��� = Get���ղ���_����������("ҽԺ�ȼ�")
        '�����ı�׼�۸�Ϊ����ҽԺ������޼ۣ�����ҽԺ������޼��ڴ˻����Ͽ����ϸ�10%��һ��ҽԺ������޼��ڴ˻������µ�5%
        
        Select Case strҽԺ���
            Case "����"
                dblҽ�� = dblҽ�� * 1.1
            Case "һ��"
                dblҽ�� = dblҽ�� * 0.95
        End Select
    End If
    
    If bln�ؼ� = True And dbl�ؼ� > dblҽ�� Then
        '����ʹ���ؼ�
        dblҽ�� = dbl�ؼ�
    End If
    
    If dblҽԺ > dblҽ�� Then
        If MsgBox("ҽԺ����" & Format(dblҽԺ, "0.000") & " ����ҽ�����ĺ�׼�ļ۸�" & Format(dblҽ��, "0.000") & "���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    �۸��ж�_���������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���¼���_����������(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim intҵ������ As Integer
    Dim lng����ID As Long
    Dim StrInput As String
    Dim str����֢ As String, str�������� As String
    Dim str��ˮ�� As String, str���� As String, strҽ�� As String, str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    
    '��ò��˳�Ժ���ּ�����֢
    gstrSQL = " Select B.���� ���ֱ���,A.����֢,A.ҵ������,A.��ˮ�� From �����ʻ� A,���ղ��� B " & _
              " Where A.����ID=B.ID And A.����=[1] ANd A.����=B.���� And A.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˳�Ժ���ּ�����֢", TYPE_����������, lng����ID)
    str�������� = Nvl(rsTemp!���ֱ���)
    str����֢ = Nvl(rsTemp!����֢)
    intҵ������ = Nvl(rsTemp!ҵ������)
    str��ˮ�� = Nvl(rsTemp!��ˮ��)
    
    'ȡ��Ժ�����Ϣ
    gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.���� ����,A.��Ժ��ʽ,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
    str����ʱ�� = Format(rsTemp!��Ժ����, "yyyyMMdd") & " 00:00:00"
    str���� = rsTemp!����
    strҽ�� = Nvl(rsTemp!ҽ��, "����")
    
    '�ò���Ա�޸Ĳ�����Ϣ�Ͳ���֢
    If frm����ѡ��_����������.ShowSelect(frmParent, intҵ������, str��������, str����֢) = False Then
        Exit Function
    End If
    
    '���ݲ��ֱ���ȡ�ò��ֵ�ID
    gstrSQL = "Select ID From ���ղ��� Where ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", TYPE_����������, str��������)
    lng����ID = rsTemp!ID
    
    '���±����ʻ�
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'����ID','" & lng����ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���������� & ",'����֢','''" & str����֢ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
    
    '���²������״̬
    StrInput = str��ˮ�� & gstrSplit_Col_���������� & intҵ������ & gstrSplit_Col_���������� & _
            str���� & gstrSplit_Col_���������� & strҽ�� & gstrSplit_Col_���������� & _
            str����ʱ�� & gstrSplit_Col_���������� & str�������� & gstrSplit_Col_���������� & _
            "1" & gstrSplit_Col_���������� & "" & gstrSplit_Col_���������� & _
            "" & gstrSplit_Col_���������� & "" & gstrSplit_Col_���������� & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_���������� & ToVarchar(str����֢, 50)
    Call ���ýӿ�_׼��_����������("09", StrInput)
    If Not ���ýӿ�_����������() Then Exit Function
    
    gcnOracle.CommitTrans
    ���¼���_���������� = True
    
    '���ܴ�����Ϣ�����˱仯����Ҫ���»�ȡ
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1   string  18      ����/סԺ��ˮ��
'    2   string  20      ���ֱ��루��Ժ��ϣ�
'    3   datetime        ��  ��Ժ����
'    4   string          ������ҽ�ƴ�����Ϣ�ļ������·�����ļ���
    StrInput = str��ˮ�� & gstrSplit_Col_���������� & str�������� & gstrSplit_Col_���������� & _
            str����ʱ�� & gstrSplit_Col_���������� & GetFileName(������Ϣ)
    Call ���ýӿ�_׼��_����������("23", StrInput)
    If Not ���ýӿ�_����������() Then
        MsgBox "��ȡ������Ϣʱ��������", vbInformation, gstrSysName
        Exit Function
    End If
    If Not AnalyFile_Deal(True) Then
        MsgBox "����������Ϣ�ļ�ʱ��������", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Public Sub TestVerifyItem()
    Dim StrInput As String, strReturn As String
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    
    Const int��ˮ�� As Integer = 0
    Const int������ˮ�� As Integer = 1
    Const int���շ�������� As Integer = 2
    Const int������־ As Integer = 3
    Const int���Ը���� As Integer = 4
    'δ��������Ŀ�ٴε��ô�����ϸ����,�����淵�ؽ��
    gstrSQL = " Select ��ˮ��,������ˮ�� From �м��_������ϸ" & _
              " Where ������־ = '0' And ��ˮ��='" & gComInfo_����������.������ˮ�� & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����������
    End With
    
    Do While Not rsTemp.EOF
        StrInput = rsTemp!��ˮ�� & gstrSplit_Col_���������� & rsTemp!������ˮ��
        Call ���ýӿ�_׼��_����������("11", StrInput)
        If ���ýӿ�_���������� Then
            strReturn = gstrReturn_����������
            '��ˮ��,������ˮ��,���շ��������,������־,���Ը����
            arrData = Split(strReturn, gstrSplit_Col_����������)
            gcn����������.Execute "zl_�м��_������ϸ_UPDATE(" & _
                            "'" & arrData(int��ˮ��) & "','" & arrData(int������ˮ��) & "'," & _
                            "'" & arrData(int���շ��������) & "','" & arrData(int������־) & "'," & _
                            "'" & arrData(int���Ը����) & "')", , adCmdStoredProc
        End If
        rsTemp.MoveNext
    Loop
End Sub

Public Function CheckItem() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '����Ƿ����δͨ�������ĸ��շ���Ŀ
    gstrSQL = " Select ��ˮ��,������ˮ��,������־ From �м��_������ϸ" & _
              " Where ������־ = '0' And ��ˮ��='" & gComInfo_����������.������ˮ�� & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����������
        CheckItem = (.RecordCount = 0)
    End With
End Function

Public Sub �˶Է��ý���_����������()
    '���м���б�������ݺ����Ķ˵����ݽ��жԱ�
    Dim StrInput As String, strReturn As String
    Dim strStart As String, strEnd As String
    Dim str������ˮ��_��ʼ As String, str������ˮ��_���� As String
    Dim cur����ͳ��_ҽ�� As Currency, cur��ͳ��_ҽ�� As Currency, cur����Ա����_ҽ�� As Currency
    Dim cur����ͳ��_ҽԺ As Currency, cur��ͳ��_ҽԺ As Currency, cur����Ա����_ҽԺ As Currency
    Dim arrReturn
    Dim rsTemp As New ADODB.Recordset
    
    If frm���ڷ�Χ_����.Show_ME(strStart, strEnd) = False Then Exit Sub
    gstrSQL = " Select min(��ˮ��) ��ʼ��ˮ��,max(��ˮ��) ������ˮ�� From " & mstrOwner & ".�м��_������Ϣ" & _
            " Where ����ʱ�� Between [1] And [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ʱ�䷶Χ�ڵ���ˮ��", CDate(strStart), CDate(strEnd))
    str������ˮ��_��ʼ = Nvl(rsTemp!��ʼ��ˮ��)
    str������ˮ��_���� = Nvl(rsTemp!������ˮ��)
    If str������ˮ��_���� = "" And str������ˮ��_��ʼ = "" Then
        MsgBox "ָ�����ڼ���û�з����κ�ҽ�����ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '׼�����ú˶Է��ýӿ�
    If Not ҽ����ʼ��_���������� Then Exit Sub
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  3       ͳ������
'    2.      string  14      ����ҽ�ƻ������
'    3.      string  18      ��Ҫ�˶Եķ��ý�����Ϣ����ʼ����/סԺ��ˮ��
'    4.      string  18      ��Ҫ�˶Եķ��ý�����Ϣ�Ľ�ֹ����/סԺ��ˮ��
'    5.      string  18      ���㽻����ˮ��
'    6.      string  18      ������¼���(����״̬��ˮ��)
'    OutputString
'    ���    ��������    ����    ����    ˵��
'    1.      number  15      �ں˶Է�Χ�ڵ�������Ϣ�ļ�¼����
'    2.      number  14  2   �ں˶Է�Χ�ڵ����м�¼�ĸ����Ը��ܶ��ۼ�ֵ
'    3.      number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ļ���ҽ��ͳ��֧���ܶ��ۼ�ֵ
'    4.      number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ĺ���Ա�����ܶ��ۼ�ֵ
'    5.      number  14  2   �ں˶Է�Χ�ڵ����м�¼�Ĵ�������ܶ��ۼ�ֵ
    StrInput = "" & gstrSplit_Col_���������� & gComInfo_����������.ҽԺ���� & gstrSplit_Col_���������� & _
               str������ˮ��_��ʼ & gstrSplit_Col_���������� & str������ˮ��_���� & gstrSplit_Col_���������� & _
               "" & gstrSplit_Col_���������� & ""
    Call ���ýӿ�_׼��_����������("16", StrInput)
    If Not ���ýӿ�_���������� Then Exit Sub
    strReturn = gstrReturn_����������
    
    '�ֽⷵ�ش�
    arrReturn = Split(strReturn, gstrSplit_Col_����������)
    cur����ͳ��_ҽ�� = Val(arrReturn(2))
    cur����Ա����_ҽ�� = Val(arrReturn(3))
    cur��ͳ��_ҽ�� = Val(arrReturn(4))
    
    '��ȡ�м���б���Ľ�����Ϣ
    gstrSQL = " Select Sum(Nvl(���λ���ͳ����,0)) ����ͳ��,SUM(Nvl(��֧�����,0)) ��ͳ��," & _
              " Sum(Nvl(���Ը����ֹ���Ա����,0)+Nvl(��ʷ���Ը�����Ա����,0)+Nvl(�����¹���Ա����,0)+Nvl(������ͨ���﹫��Ա����,0)+Nvl(�ֶ��Ը�����Ա����,0)+Nvl(תԺ�������빫��Ա,0)) ����Ա����" & _
              " From " & mstrOwner & ".�м��_������Ϣ" & _
              " Where ��ˮ��>=[1] And ��ˮ��<=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�м���еĽ�����Ϣ", str������ˮ��_��ʼ, str������ˮ��_����)
    cur����ͳ��_ҽԺ = Val(Nvl(rsTemp!����ͳ��, 0))
    cur��ͳ��_ҽԺ = Val(Nvl(rsTemp!��ͳ��, 0))
    cur����Ա����_ҽԺ = Val(Nvl(rsTemp!����Ա����, 0))
    
    '����Ƿ���ͬ
    If Format(cur����ͳ��_ҽ��, "#####0.00") = Format(cur����ͳ��_ҽԺ, "#####0.00") And _
    Format(cur��ͳ��_ҽ��, "#####0.00") = Format(cur��ͳ��_ҽԺ, "#####0.00") And _
    Format(cur����Ա����_ҽ��, "#####0.00") = Format(cur����Ա����_ҽԺ, "#####0.00") Then
        MsgBox "����ͳ��󲡲���������Ա�������������һ�£�", vbInformation, gstrSysName
    Else
        MsgBox "�˶Խ�һ�£�" & vbCrLf & _
        "����ͳ���ҽ����" & Format(cur����ͳ��_ҽ��, "#####0.00") & Space(10) & "��ҽԺ��" & Format(cur����ͳ��_ҽԺ, "#####0.00") & vbCrLf & _
        "�󲡲�������ҽ����" & Format(cur��ͳ��_ҽ��, "#####0.00") & Space(10) & "��ҽԺ��" & Format(cur��ͳ��_ҽԺ, "#####0.00") & vbCrLf & _
        "����Ա��������ҽ����" & Format(cur����Ա����_ҽ��, "#####0.00") & Space(10) & "��ҽԺ��" & Format(cur����Ա����_ҽԺ, "#####0.00"), vbInformation, gstrSysName
    End If
End Sub

Private Function GetSequence(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�շ�ϸĿID As Long) As String
    '���ȡһ��������¼����ˮ�ţ����ڸ������ʣ�
    Dim rsTemp As New ADODB.Recordset
    GetSequence = ""
    
    gstrSQL = " Select NO,��¼����,��¼״̬,��� From סԺ���ü�¼" & _
              " Where �շ�ϸĿID=[1] And ����ID=[2] And ��ҳID=[3]" & _
              " And ��¼״̬=1 And Nvl(ʵ�ս��,0)>0 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ˮ��", lng�շ�ϸĿID, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        GetSequence = rsTemp!NO & String(3 - Len(rsTemp!��¼����), "0") & rsTemp!��¼���� & _
                    String(3 - Len(rsTemp!��¼״̬), "0") & rsTemp!��¼״̬ & _
                    String(3 - Len(rsTemp!���), "0") & rsTemp!���
    Else
        Call DebugTool("δ�ҵ�ԭʼ������ϸ[����ID:" & lng����ID & "|��ҳID:" & lng��ҳID & "|�շ�ϸĿID:" & lng�շ�ϸĿID)
    End If
End Function

Public Function תΪ��ͨ����_����(ByVal lng����ID As Long) As Boolean
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡԭ������ˮ��
    gstrSQL = "Select ��ˮ�� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ˮ��", TYPE_����������, lng����ID)
    gComInfo_����������.������ˮ�� = rsTemp!��ˮ��
    
    '���þ���Ǽ����Ͻӿ�
'    InputString
'    ���    ��������    ����    ����    ˵��
'    1.      string  18      ����/סԺ��ˮ��
    Call ���ýӿ�_׼��_����������("15", gComInfo_����������.������ˮ��)
    If Not ���ýӿ�_����������() Then Exit Function
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    תΪ��ͨ����_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AnalyBalance(ByVal strBalance As String)
    Dim arrBalance
    Dim STRNAME As String
    Dim dblMoney As Double
    Dim intDO As Integer, intCOUNT As Integer
    '�������㷵�ش�������Ϣ��䵽�ṹ����
    
    pre_Balance.cur�����ʻ� = 0
    pre_Balance.curҽ������ = 0
    pre_Balance.cur����Ա���� = 0
    pre_Balance.cur�󲡻��� = 0
    
    arrBalance = Split(strBalance, "|")
    intCOUNT = UBound(arrBalance)
    For intDO = 0 To intCOUNT
        STRNAME = Split(arrBalance(intDO), ";")(0)
        dblMoney = Val(Split(arrBalance(intDO), ";")(1))
        Select Case STRNAME
        Case "�����ʻ�"
            pre_Balance.cur�����ʻ� = dblMoney
        Case "ҽ������"
            pre_Balance.curҽ������ = dblMoney
        Case "����Ա��������"
            pre_Balance.cur����Ա���� = dblMoney
        Case "�󲡻���"
            pre_Balance.cur�󲡻��� = dblMoney
        End Select
    Next
End Sub
