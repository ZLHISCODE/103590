Attribute VB_Name = "mdl�Ͼ���"
Option Explicit
Private mstrPatID As String
Private mobjSystem As New FileSystemObject
Private mobjStream As TextStream
Private mcur������� As Currency
Public gstr��ȷ���� As String
Public gstr��Ա��� As String
'Private mdomInput As New MSXML2.DOMDocument
'Private mdomOutput As New MSXML2.DOMDocument
Private mblnInit As Boolean
Public gblnBill As Boolean          'ҽ��Ʊ�ݿ���
Public gblnCancel_�Ͼ� As Boolean     '�������Ԥ��ҽ����������������ȡ��
Public gintBills As Integer
Public glng����ID As Long
Public glng����ID As Long
Public gint��ϸ�� As Integer
Public gint�վݷ�Ŀ As Integer
Public gcnNJSYB As New ADODB.Connection

Private Type patInfo_�Ͼ���
    ҽ���� As String
    ����ʱ�� As String
    �������� As String
    ҽ������ As String
    ҽ������ As String
    ���ֱ��� As String
    �������� As String
    ҽ����������� As String
    ҽ����������� As String
    �����˱��� As String
End Type
Public gPatInfo_�Ͼ��� As patInfo_�Ͼ���

Private Type detailFee_�Ͼ���
    �к� As Double
    ҽ���� As String
    סԺ��� As String
    �������� As String
    ��־ As String
    ���÷���ʱ�� As String
    ҽԺ���� As String
    ҽԺ�Ա���  As String
    ҽ������ As String
    ���� As String
    ������λ As String
    ���� As Double
    ���� As Double
    �����˱��� As String
    ���������� As String
    ���� As String
    �������� As String
    ��� As String
End Type
Private mDetailFee_�Ͼ��� As detailFee_�Ͼ���

Private Type feeBalance_�Ͼ���
    סԺ��� As String
    ҽ������ As String
    ���÷���ʱ�� As String
    ������úϼ� As Double
    ҩ�Ѻϼ� As Double
    ������Ŀ�ϼ� As Double
    ������� As Double
    ҽ����Χ���� As Double
    �����ʻ�֧�� As Double
    ͳ��֧�� As Double
    ��֧�� As Double
    �����Ը� As Double
    �ڳ������ʻ� As Double
    ��ĩ�����ʻ� As Double
    ����Ա���� As String
    ���ݺ� As String
    ���� As String
    �Ż�1 As Double
    �Ż�2 As Double
    �Ż�3 As Double
    �ͱ��ʻ�֧�� As Double
End Type
Public mFeeBalance As feeBalance_�Ͼ���

Public gstr��ע As String

Public Function ҽ����ʼ��_�Ͼ���() As Boolean
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gblnBill = (GetSetting("ZLSOFT", "����ģ��\ҽ��Ʊ�ݹ���", "ҽ��Ʊ�ݹ���", 0) = 1)
    
    If Not mblnInit And gblnBill Then
        strServer = AnalyServer(gcnOracle.ConnectionString)
        Call AnalyConf(strUser, strPass, strServer)
        With gcnNJSYB
            If .State = 1 Then .Close
            .Provider = "MsDataShape"
            .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
            If Err <> 0 Then
                MsgBox "�����м��û�ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End With
        
        '��ȡ��ϸ�����վݷ�Ŀ����
        gstrSQL = " Select ��ϸ��,�վݷ�Ŀ From Ʊ�ݴ�ӡ����"
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        If rsTemp.RecordCount <> 0 Then
            gint��ϸ�� = Nvl(rsTemp!��ϸ��, 0)
            gint�վݷ�Ŀ = Nvl(rsTemp!�վݷ�Ŀ, 0)
        End If
        
        If gint��ϸ�� = 0 Or gint�վݷ�Ŀ = 0 Then
            MsgBox "��ʹ��ҽ��Ʊ�ݹ���������Ʊ�ݴ�ӡ������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    mblnInit = True
     ҽ����ʼ��_�Ͼ��� = True
End Function

Public Function ��ݱ�ʶ_�Ͼ���(Optional bytType As Byte, Optional lng����ID As Long) As String
    
    On Error GoTo errorhandle
    If bytType = 0 Or bytType = 3 Then
        If gblnBill Then
            '����Ƿ������û���Ʊ��,û����������������֤;��סԺ�ڽ���ʱ��ʹ��Ʊ��,���Բ����
            glng����ID = GetSetting("ZLSOFT", "����ģ��\ҽ��Ʊ�ݹ���\����", "�����շ�Ʊ������", 0)
            glng����ID = GetInvoiceGroupID(1, 1, glng����ID, glng����ID)
            If glng����ID <= 0 Then
                Select Case glng����ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õ�ҽ���շ�Ʊ��,��������һ��ҽ��Ʊ�ݻ����ñ��ع���ҽ��Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���ع��õ�ҽ��Ʊ���Ѿ�����,��������һ��ҽ��Ʊ�ݻ��������ñ��ع���ҽ��Ʊ�ݣ�", vbInformation, gstrSysName
                End Select
                Exit Function
            End If
        End If
        ��ݱ�ʶ_�Ͼ��� = frmIdentify�Ͼ���.Identify(bytType, lng����ID)
    Else
        ��ݱ�ʶ_�Ͼ��� = frm���ݽ���.getFeeBalance(bytType, lng����ID)
        Unload frm���ݽ���
    End If
    
    gblnCancel_�Ͼ� = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_�Ͼ���(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional strAdvance As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '�ֶΣ�������,����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��

    Dim rsTemp As New ADODB.Recordset, curCount As Currency, dbl�ֽ� As Double
    Dim strFile As String, strWrite As String
    Dim strTemp As String
    Dim intOrder As Integer
    Dim dblʵ�ս�� As Double, str��ע As String
    Dim intSubInsure As Integer, strYHLB As String, dblSubBalance As Double, strSubInsureNO As String, intSubDisable As Integer  '��ҽ����ţ���ҽ���ʻ�����ҽ���ţ�ͣ�ñ�־
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
'    Dim InvokeServer As String
    'ɾ�����ܴ��ڵ�ǰ�ν�����Ϣ�ļ�
    On Error Resume Next
    Call Kill("C:\NJYB\mzjshz.xml")
    
    On Error GoTo errorhandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ü�¼�����ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡ��ǰ������ҽ�������Ϣ(�������|�Ż����|ҽ����|���|ͣ��)
    gstrSQL = " Select ����֤��||'||||' AS ����֤�� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ������ҽ�������Ϣ", TYPE_�Ͼ���, CLng(rs��ϸ!����ID))
    intSubInsure = Val(Split(rsTemp!����֤��, "|")(0))
    strYHLB = Split(rsTemp!����֤��, "|")(1)
    strSubInsureNO = Split(rsTemp!����֤��, "|")(2)
    dblSubBalance = Val(Split(rsTemp!����֤��, "|")(3))
    intSubDisable = Val(Split(rsTemp!����֤��, "|")(4)) 'ͣ��������ʹ��ͳ��֧������˼����ϸ�޴���
    Select Case strYHLB
    Case "����"
        strYHLB = "1"
    Case "����"
        strYHLB = "2"
    Case "�����"
        strYHLB = "3"
    Case Else
        strYHLB = "0"
    End Select
    
    curCount = 0
    While Not rs��ϸ.EOF
         curCount = curCount + rs��ϸ!ʵ�ս��
        rs��ϸ.MoveNext
    Wend
    rs��ϸ.MoveFirst
    
    strAdvance = ""
    If gblnBill Then
        '���Ʊ��
        gintBills = AnalyBill(rs��ϸ)
        If IsEnough() = False Then
            MsgBox "�����շѽ�ʹ��" & gintBills & "��ҽ��Ʊ�ݣ�����ǰʣ���������㣬�����Ʊ�ݺ������շѣ�", vbInformation, gstrSysName
            Exit Function
        End If
        strAdvance = GetNextBill(glng����ID) '��ǰ̨���򷵻ر��ν�����ʹ�õ�Ʊ�ݿ�ʼ����
    End If
    
    'ȡ��������Ϣ��������
    mstrPatID = rs��ϸ!����ID
    With gPatInfo_�Ͼ���
        .����ʱ�� = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss")        '�õ�����ʱ��
        .ҽ������ = Nvl(rs��ϸ!������)                                               '�õ�ҽ������
    End With
    
    If Trim(gPatInfo_�Ͼ���.ҽ������) = "" Then
        MsgBox "ҽ�������շѱ�������ҽ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select A.�����ʼ� as ҽ������,decode(c.λ��,null,c.����,c.λ��) as ҽ�����ұ���,C.���� as ҽ���������� from ��Ա�� A,������Ա B,���ű� C,�ٴ����� D " & _
              "where A.id=B.��Աid and B.����id = C.id and C.id=D.����id and B.ȱʡ=1 and  A.����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ������", CStr(rs��ϸ!������))
    If rsTemp.EOF Then
        MsgBox "δ��Ӧ���ҵ����ƿ�Ŀ����,������ȷ��Ӧ", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gPatInfo_�Ͼ���
        .ҽ������ = rsTemp!ҽ������                                               'ȡ��ҽ������
        .ҽ����������� = rsTemp!ҽ�����ұ���
        .ҽ����������� = rsTemp!ҽ����������
        .�����˱��� = UserInfo.���
    End With
    
    If InitXML = False Then Exit Function
    Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
    Call InsertChild(nodRow, "TBR", gPatInfo_�Ͼ���.ҽ����)
    Call InsertChild(nodRow, "XM", gPatInfo_�Ͼ���.��������)
    Call InsertChild(nodRow, "YSM", gPatInfo_�Ͼ���.ҽ������)
    Call InsertChild(nodRow, "YSXM", gPatInfo_�Ͼ���.ҽ������)
    Call InsertChild(nodRow, "BZBM", gPatInfo_�Ͼ���.���ֱ���)
    Call InsertChild(nodRow, "KSM", gPatInfo_�Ͼ���.ҽ�����������)
    Call InsertChild(nodRow, "KSMC", gPatInfo_�Ͼ���.ҽ�����������)
        
    mdomInput.Save "C:\NJYB\mzjzxx.xml"
    
    'ȡ����ϸ������������
    gstrSQL = "select ҽԺ���� from ������� where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽԺ����", TYPE_�Ͼ���)
    If rsTemp.EOF Then
        MsgBox "ҽԺ����δ����,��������ҽԺ����", vbInformation, gstrSysName
        Exit Function
    End If
    With mDetailFee_�Ͼ���
        .�������� = gPatInfo_�Ͼ���.��������
        .���÷���ʱ�� = gPatInfo_�Ͼ���.����ʱ��
        .ҽԺ���� = rsTemp!ҽԺ����
        .�����˱��� = gPatInfo_�Ͼ���.�����˱���
    End With
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ����Ŀ", TYPE_�Ͼ���, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    rs��ϸ.MoveFirst
    
    rs��ϸ.MoveFirst
    If InitXML = False Then Exit Function
    Do Until rs��ϸ.EOF
        If rs��ϸ!ʵ�ս�� <> 0 Then
            '����������ҽ���õ�������ϸ��ͳ����
            dblʵ�ս�� = rs��ϸ!ʵ�ս��
            If intSubInsure <> 0 And intSubDisable = 0 Then '�ͱ����ʻ�δͣ��
                If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
                If Not gobjInsure_Obj(intOrder).CalcSingleRecord(rs��ϸ!�շ�ϸĿID, dblʵ�ս��, gstr��ע, intSubInsure, rs��ϸ.AbsolutePosition) Then Exit Function
            End If
            
            '׼���ϴ�����
            gstrSQL = "select decode(A.���,'5',0,'6',0,'7',0,1) ��־,A.����,nvl(b.�����װ,1) �����װ,C.��Ŀ����,a.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���" & _
                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C where A.id = C.�շ�ϸĿid and A.id=B.ҩƷid(+) and A.id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ϸ��ϸ", CLng(rs��ϸ!�շ�ϸĿID))
            With mDetailFee_�Ͼ���
                .��־ = rsTemp!��־
                .���� = rsTemp!����
                .ҽ������ = rsTemp!��Ŀ����
                .������λ = Nvl(rsTemp!���㵥λ)
                .������λ = ToVarchar(.������λ, 10)
                .���� = Val(Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� / rsTemp!�����װ), "#0.0000;-#0.0000;0;"))
                .���� = rs��ϸ!���� / rsTemp!�����װ
                .���� = Nvl(rsTemp!����)
                .�������� = Nvl(rsTemp!��������)
                .��� = Nvl(rsTemp!���)
            End With
            
            Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
            Call InsertChild(nodRow, "TBR", gPatInfo_�Ͼ���.ҽ����)
            Call InsertChild(nodRow, "XM", gPatInfo_�Ͼ���.��������)
            Call InsertChild(nodRow, "BZ", mDetailFee_�Ͼ���.��־)
            Call InsertChild(nodRow, "ZBM", mDetailFee_�Ͼ���.ҽ������)
            Call InsertChild(nodRow, "SL", mDetailFee_�Ͼ���.����)
            Call InsertChild(nodRow, "DJ", mDetailFee_�Ͼ���.����)
            Call InsertChild(nodRow, "YHLB", strYHLB)
            Call InsertChild(nodRow, "YHJ", Val(Format(dblʵ�ս�� / (rs��ϸ!���� / rsTemp!�����װ), "#0.0000;-#0.0000;0;")))
        End If
          
        rs��ϸ.MoveNext
     Loop

    mdomInput.Save "C:\NJYB\mzcfsj.xml"
    Call DebugTool("��ϸ�ļ��Ѳ���")
    
    '����ҽ��������
    strTemp = frm���ݽ���.getFeeBalance
    On Error Resume Next
    Unload frm���ݽ���
    On Error GoTo errorhandle
    If strTemp = "" Then
        MsgBox "��ȡҽ�������ļ����̱���ֹ,�޷����Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ����ϢΪ���������׼��
    '�����Ĵ���
    If InitXML = False Then Exit Function
    Set mdomInput = New MSXML2.DOMDocument
    If mdomInput.Load("c:\njyb\mzjshz.xml") = False Then
        MsgBox "ҽ������������ֵ��ʽ����ȷ��", vbInformation, gstrSysName
    Else
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        With mFeeBalance
            .ҽ������ = nodRowset.selectSingleNode("TBR").Text
            .������úϼ� = nodRowset.selectSingleNode("ZFY").Text
            .������� = nodRowset.selectSingleNode("GRZL").Text
            .�����ʻ�֧�� = nodRowset.selectSingleNode("ZHZF").Text
            .ͳ��֧�� = nodRowset.selectSingleNode("YBZF").Text
            .�����Ը� = nodRowset.selectSingleNode("GRZF").Text
            .���ݺ� = nodRowset.selectSingleNode("DJH").Text
            If nodRowset.selectSingleNode("FYLB").Text = "�ž�" Then
                .��֧�� = nodRowset.selectSingleNode("ZFY").Text
            Else
                .��֧�� = 0
            End If
            .���� = nodRowset.selectSingleNode("XZMC").Text
            .�Ż�1 = Val(nodRowset.selectSingleNode("YH1").Text)
            .�Ż�2 = Val(nodRowset.selectSingleNode("YH2").Text)
            .�Ż�3 = Val(nodRowset.selectSingleNode("YH3").Text)
        End With
    End If
    Call DebugTool("��ɸ������ݵĶ�ȡ")
    If curCount <> CCur(mFeeBalance.������úϼ�) Then
        MsgBox "��ע�⣺ҽ�����ط��úϼ���ҽԺ������úϼƲ���" & vbCrLf & _
            "ҽԺ��" & curCount & Space(10) & "ҽ����" & mFeeBalance.������úϼ�
    End If
    mcur������� = nodRowset.selectSingleNode("ZHYE").Text
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mstrPatID & "," & TYPE_�Ͼ��� & ",'�ʻ����','" & mcur������� & "')"
    Call DebugTool("�����ʻ����:" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

'����
    gstr��Ա��� = nodRowset.selectSingleNode("XZMC").Text
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mstrPatID & "," & TYPE_�Ͼ��� & ",'��Ա���',''" & gstr��Ա��� & "'')"
    Call DebugTool("������Ա���:" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    Call DebugTool("���ؽ��㷽ʽ")
    str���㷽ʽ = "�����ʻ�;" & mFeeBalance.�����ʻ�֧�� & ";0"
    If mFeeBalance.ͳ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & mFeeBalance.ͳ��֧�� & ";0"
    End If
    If mFeeBalance.��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|��ͳ��;" & mFeeBalance.��֧�� & ";0"
    End If
    If mFeeBalance.�Ż�1 <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|������;" & mFeeBalance.�Ż�1 & ";0"
    End If
    If mFeeBalance.�Ż�2 <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|���Ƽ���;" & mFeeBalance.�Ż�2 & ";0"
    End If
    If mFeeBalance.�Ż�3 <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|������Ż�;" & mFeeBalance.�Ż�3 & ";0"
    End If
    '����ͱ����ʻ�֧�����
    dbl�ֽ� = curCount - mFeeBalance.�����ʻ�֧�� - mFeeBalance.ͳ��֧�� - mFeeBalance.��֧�� - mFeeBalance.�Ż�1 - mFeeBalance.�Ż�2 - mFeeBalance.�Ż�3
    If dbl�ֽ� >= dblSubBalance Then
        mFeeBalance.�ͱ��ʻ�֧�� = dblSubBalance
    Else
        mFeeBalance.�ͱ��ʻ�֧�� = dbl�ֽ�
    End If
    str���㷽ʽ = str���㷽ʽ & "|�ͱ��ʻ�;" & mFeeBalance.�ͱ��ʻ�֧�� & ";0"
    
    gblnCancel_�Ͼ� = False '��ʱ������ȡ������
    �����������_�Ͼ��� = True
    Exit Function
errorhandle:
    Call DebugTool("�������ʱ��������" & Err.Description)
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�Ͼ���(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim intDO As Integer
    Dim lng����ID As Long
    Dim intOrder As Integer
    Dim strNO As String
    Dim strRecord As String         '��¼ʹ�õķ�Ʊ�嵥
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim intSubInsure As Integer, strYHLB As String, dblSubBalance As Double, strSubInsureNO As String, intSubDisable As Integer  '��ҽ����ţ���ҽ���ʻ�����ҽ���ţ�ͣ�ñ�־
    Dim str���׺� As String, str���1 As String, str���2 As String, str���� As String
    On Error GoTo errorhandle
    
    gstrSQL = " Select ����ID From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    '��ȡ��ǰ������ҽ�������Ϣ(�������|�Ż����|ҽ����|���|ͣ��)
    gstrSQL = " Select ����֤��||'||||' AS ����֤�� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ������ҽ�������Ϣ", TYPE_�Ͼ���, lng����ID)
    intSubInsure = Val(Split(rsTemp!����֤��, "|")(0))
    strYHLB = Split(rsTemp!����֤��, "|")(1)
    strSubInsureNO = Split(rsTemp!����֤��, "|")(2)
    dblSubBalance = Val(Split(rsTemp!����֤��, "|")(3))
    intSubDisable = Val(Split(rsTemp!����֤��, "|")(4))
    Select Case strYHLB
    Case "����"
        strYHLB = "1"
    Case "����"
        strYHLB = "2"
    Case "�����"
        strYHLB = "3"
    Case Else
        strYHLB = "0"
    End Select
    '��ɵͱ�����
    If intSubInsure <> 0 Then
        If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
        
        str���׺� = "05"
        str���1 = strSubInsureNO & "|" & ToVarchar(gstr��λ����, 50) & "|" & lng����ID & "|" & mFeeBalance.������úϼ� & _
                  "|" & 0 & "|" & mFeeBalance.�ͱ��ʻ�֧�� & "|" & UserInfo.���� & "|" & gstr��ע
        If Not gobjInsure_Obj(intOrder).CallAPI(str���׺�, str���1, str���2, str����) Then Exit Function
    End If
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͼ��� & "," & mstrPatID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.������úϼ� & "," & mFeeBalance.������� + mFeeBalance.�����Ը� & ",0," & _
              mFeeBalance.ҽ����Χ���� & "," & mFeeBalance.ͳ��֧�� & "," & mFeeBalance.��֧�� & "," & _
              "0," & mFeeBalance.�����ʻ�֧�� & ",'" & mFeeBalance.���ݺ� & "',null,null,'" & gPatInfo_�Ͼ���.���ֱ��� & "|" & gPatInfo_�Ͼ���.�������� & IIf(str���� = "", "", "|" & Split(str����, "|")(0) & "|" & intSubInsure) & "|" & mFeeBalance.ҽ������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Ͼ���ҽ��")
    
    '����Ʊ��ʹ�ü�¼
'    zl_Ʊ��ʹ����ϸ_Insert (
'    ����ID_IN IN Ʊ��ʹ����ϸ.����ID%TYPE,
'    Ʊ��_IN IN Ʊ��ʹ����ϸ.Ʊ��%TYPE,
'    ����_IN IN Ʊ��ʹ����ϸ.����%TYPE,
'    ����_IN IN Ʊ��ʹ����ϸ.����%TYPE,
'    ԭ��_IN IN Ʊ��ʹ����ϸ.ԭ��%TYPE,
'    ����ID_IN IN Ʊ��ʹ����ϸ.����ID%TYPE,
'    ʹ��ʱ��_IN IN Ʊ��ʹ����ϸ.ʹ��ʱ��%TYPE,
'    ʹ����_IN IN Ʊ��ʹ����ϸ.ʹ����%TYPE
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        For intDO = 1 To gintBills
            strNO = GetNextBill(glng����ID)
            gstrSQL = "zl_Ʊ��ʹ����ϸ_Insert(" & glng����ID & ",1,'" & strNO & "',1,1," & lng����ID & ",sysdate,'" & UserInfo.���� & "')"
            gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
            strRecord = strRecord & "," & strNO
        Next
        gcnNJSYB.CommitTrans
        blnTrans = False
        
        If strRecord <> "" Then
            strRecord = Mid(strRecord, 2)
            Err.Raise 9000, gstrSysName, "����ҽ��ʹ��Ʊ�ݺţ�" & strRecord
        End If
    End If
    
    gblnCancel_�Ͼ� = True
    �������_�Ͼ��� = True
    Exit Function
    
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_�Ͼ���(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    Dim intOrder As Integer, intSubInsure As Integer, strSub˳��� As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lng����ID As Long
    Dim str���׺� As String, str���1 As String, str���2 As String, str���� As String
    On Error GoTo errorhandle
    
    gstrSQL = "select distinct A.����id  from ������ü�¼ A,������ü�¼ B where A.��¼״̬=2 and A.NO=B.NO and B.����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����id", lng����ID)
    lng����ID = rsTemp!����ID
    
    gstrSQL = "select * from ���ս����¼ where ��¼id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ԭʼ��¼", lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "���ս����¼��ԭʼ���ʵ��ݲ�����,�������˷�"
        Exit Function
    Else
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ͼ��� & "," & rsTemp!����ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!�������ý�� & "," & -rsTemp!ȫ�Ը���� & "," & -rsTemp!�����Ը���� & "," & -rsTemp!����ͳ���� & "," & -rsTemp!ͳ�ﱨ����� & "," & -rsTemp!���Ը���� & "," & _
              "0," & -rsTemp!�����ʻ�֧�� & ",null,null,null,null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���ʼ�¼")
    End If
    
    If InStr(1, rsTemp!֧��˳���, "|") <> 0 Then
        intSubInsure = Val(Split(rsTemp!֧��˳���, "|")(2))
        strSub˳��� = Split(rsTemp!֧��˳���, "|")(1)
        If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
        
        str���׺� = "06"
        str���1 = strSub˳��� & "|" & UserInfo.����
        If Not gobjInsure_Obj(intOrder).CallAPI(str���׺�, str���1, str���2, str����) Then Exit Function
    End If
    
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        '���������ջط�Ʊ��¼
        gstrSQL = " Select * From Ʊ��ʹ����ϸ Where ����ID=" & lng����ID
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        With rsTemp
            Do While Not .EOF
                gstrSQL = "zl_Ʊ��ʹ����ϸ_Insert(" & !����ID & ",1,'" & !���� & "',2,2," & lng����ID & ",sysdate,'" & UserInfo.���� & "')"
                gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
                .MoveNext
            Loop
        End With
        gcnNJSYB.CommitTrans
        blnTrans = False
    End If
    
    ����������_�Ͼ��� = True
    Exit Function
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_�Ͼ���(rsExse As Recordset, ByVal lng����ID As Long, Optional strAdvance As String) As String
    Dim bytType As Byte
    Dim strFile As String, strWrite As String
    Dim strStream As String
    Dim dblSettleSum As Double
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim a As Double
    'ɾ�����ܴ��ڵ�ǰ�ν�����Ϣ�ļ�
    On Error Resume Next
    
    strAdvance = ""
    If gblnBill Then
        glng����ID = GetSetting("ZLSOFT", "����ģ��\ҽ��Ʊ�ݹ���\סԺ", "�����շ�Ʊ������", 0)
        glng����ID = GetInvoiceGroupID(3, 1, glng����ID, glng����ID)
        If glng����ID <= 0 Then
            Select Case glng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�ҽ���շ�Ʊ��,��������һ��ҽ��Ʊ�ݻ����ñ��ع���ҽ��Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���ع��õ�ҽ��Ʊ���Ѿ�����,��������һ��ҽ��Ʊ�ݻ��������ñ��ع���ҽ��Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        strAdvance = GetNextBill(glng����ID) '��ǰ̨���򷵻ر��ν�����ʹ�õ�Ʊ�ݿ�ʼ����
    End If
    
    Call Kill("C:\NJYB\CYJSD.XML")
    On Error GoTo errorhandle
    '�ϴ���δ�ϴ�����ϸ����
    gstrSQL = "select ˳��� from �����ʻ� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "˳���", lng����ID)
    mDetailFee_�Ͼ���.סԺ��� = rsTemp!˳���
    
    'δ������Ŀ
    gstrSQL = "select b.���� as ��Ŀ from סԺ���ü�¼ a ,�շ���ĿĿ¼ b  where a.�շ�ϸĿid=b.id and " & _
              "a.����id=[1] and a.��ҳid= [2] and not exists ( select 1 from ����֧����Ŀ d where d.����=[3] and a.�շ�ϸĿid=d.�շ�ϸĿid)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ϸ��ϸ", CLng(rsExse!����ID), CLng(rsExse!��ҳID), TYPE_�Ͼ���)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "��Ŀ��δ�Ա���: " & rsTemp!��Ŀ, vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ļ�
'    strFile = "C:\NJYB\ZYFYMX.XML"
'    Call writeTxtFile(strFile, "")
    If InitXML = False Then Exit Function
    Do Until rsExse.EOF
        If rsExse!�Ƿ��ϴ� = 1 Or rsExse!��� = 0 Then GoTo haddeliver           '�ҳ����ϴ���¼
        gstrSQL = "select Rownum ���,decode(A.���,'5',0,'6',0,'7',0,1) ��־,A.����,A.����,C.��Ŀ����,A.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���" & _
                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C where A.id = C.�շ�ϸĿid and A.id=B.ҩƷid(+) and A.id = [1] And C.����=[2]"
'        gstrSQL = "select Rownum ���,decode(A.���,'5',0,'6',0,'7',0,1) ��־,A.����,A.����,C.��Ŀ����,A.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���,d.�ּ�" & _
'                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C,�շѼ�Ŀ d where A.id = C.�շ�ϸĿid and a.id=d.id and A.id=B.ҩƷid(+) and A.id =" & rsExse!�շ�ϸĿID & " And C.����=" & TYPE_�Ͼ���
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ϸ��ϸ", CLng(rsExse!�շ�ϸĿID), TYPE_�Ͼ���)
        If rsTemp.RecordCount <> 0 Then
            With mDetailFee_�Ͼ���
                .�к� = rsTemp!���
                .��־ = rsTemp!��־
                .���÷���ʱ�� = Format(rsExse!�Ǽ�ʱ��, "yyyyMMdd")
                .ҽԺ�Ա��� = rsTemp!����
                .ҽ������ = rsTemp!��Ŀ����
                .���� = rsTemp!����
                .������λ = zlCommFun.Nvl(rsTemp!���㵥λ)
                .���� = Format(rsExse!��� / rsExse!����, "#0.0000;-#0.0000;0;")
'                .���� = Format(rsTemp!�ּ�, "#0.0000;-#0.0000;0;")
                .���� = rsExse!����
                .���� = zlCommFun.Nvl(rsTemp!����)
                .�������� = zlCommFun.Nvl(rsTemp!��������)
                .��� = zlCommFun.Nvl(rsTemp!���)
            End With
            
            gstrSQL = "select b.�����ʼ� as ҽ������,a.������ as ҽ�� from סԺ���ü�¼ a,��Ա�� b where a.NO=[1] and a.���=[2]" & _
                    " and mod(a.��¼����,10)=[3] and a.��¼״̬=[4] And a.������ = b.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ������", CStr(rsExse!NO), CLng(rsExse!���), CInt(rsExse!��¼����), CInt(rsExse!��¼״̬))
            If rsTemp.RecordCount = 0 Then
                'ȡסԺҽʦ
                gstrSQL = "select b.�����ʼ� as ҽ������,a.סԺҽʦ as ҽ�� from ������ҳ a,��Ա�� b  where a.סԺҽʦ=b.���� and  " & _
                          "a.����id= [1] and a.��ҳid= [2] and rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡסԺҽʦ", CLng(rsExse!����ID), CLng(rsExse!��ҳID))
            End If
            mDetailFee_�Ͼ���.�����˱��� = rsTemp!ҽ������
            mDetailFee_�Ͼ���.���������� = rsTemp!ҽ��
            
            a = a + 1
            Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
            Call InsertChild(nodRow, "ID", a)
            Call InsertChild(nodRow, "XH", mDetailFee_�Ͼ���.סԺ���)
            Call InsertChild(nodRow, "BZ", mDetailFee_�Ͼ���.��־)
            Call InsertChild(nodRow, "SJ", mDetailFee_�Ͼ���.���÷���ʱ��)
            Call InsertChild(nodRow, "ZBM", mDetailFee_�Ͼ���.ҽ������)
            Call InsertChild(nodRow, "SL", mDetailFee_�Ͼ���.����)
            Call InsertChild(nodRow, "DJ", mDetailFee_�Ͼ���.����)
            Call InsertChild(nodRow, "YSM", mDetailFee_�Ͼ���.�����˱���)
            Call InsertChild(nodRow, "YS", mDetailFee_�Ͼ���.����������)
        End If
haddeliver:
        dblSettleSum = dblSettleSum + rsExse!���           '�ó������ܽ��
        rsExse.MoveNext
    Loop
    '�ر��ļ�
    mdomInput.Save "C:\NJYB\zyfymx.xml"
'    Call writeTxtFile(strFile, "", False)
    
    bytType = 9                          '��ʾסԺԤ����״̬
    
    strStream = frm���ݽ���.getFeeBalance(bytType)
    On Error Resume Next
    Unload frm���ݽ���
    On Error GoTo errorhandle
    If strStream = "" Then
        MsgBox "��ȡҽ�������ļ����̱���ֹ,�޷����Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
'    With mFeeBalance
'        .סԺ��� = analyseStr(strStream, 1, 20)
'        .������úϼ� = Val(analyseStr(strStream, 35, 10))
'        .ҽ����Χ���� = Val(analyseStr(strStream, 65, 10))
'        .������� = Val(analyseStr(strStream, 75, 10))
'        .�����Ը� = Val(analyseStr(strStream, 85, 10))
'        .ͳ��֧�� = Val(analyseStr(strStream, 95, 10))
'        .��֧�� = Val(analyseStr(strStream, 105, 10))
'        .�����ʻ�֧�� = Val(analyseStr(strStream, 115, 10))
'    End With
    Set mdomInput = New MSXML2.DOMDocument
    If mdomInput.Load("c:\njyb\cyjsd.xml") = False Then
        MsgBox "ҽ������������ֵ��ʽ����ȷ��", vbInformation, gstrSysName
    Else
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
          With mFeeBalance
         .סԺ��� = nodRowset.selectSingleNode("XH").Text
         .������úϼ� = nodRowset.selectSingleNode("ZFY").Text
         .������� = nodRowset.selectSingleNode("GRZL").Text
         .�����Ը� = nodRowset.selectSingleNode("GRZF").Text
         .ͳ��֧�� = nodRowset.selectSingleNode("YBZF").Text
         .�����ʻ�֧�� = nodRowset.selectSingleNode("ZHZF").Text
         End With
    End If
    mcur������� = nodRowset.selectSingleNode("ZHYE").Text
    If mFeeBalance.סԺ��� <> mDetailFee_�Ͼ���.סԺ��� Then
        MsgBox "�˽��ʲ�����ҽ�������ļ��в��˲�һ��,���ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    If Format(dblSettleSum, "#0.00") <> Format(mFeeBalance.������úϼ�, "#0.00") Then
        MsgBox "��ע��:ҽԺ�ܷ�����ҽ�����ķ��ص��ܷ��ò�һ��" & vbCrLf & _
        "�ܷ���:(ҽԺ)��" & Format(dblSettleSum, "#0.00") & Space(10) & "(ҽ��)��" & Format(mFeeBalance.������úϼ�, "#0.00"), vbInformation, gstrSysName
    End If

    strStream = "ͳ�����;" & mFeeBalance.ͳ��֧�� & ";0"
    If mFeeBalance.�����ʻ�֧�� <> 0 Then
        strStream = strStream & "|�����ʻ�;" & mFeeBalance.�����ʻ�֧�� & ";0"
    End If
    If mFeeBalance.��֧�� <> 0 Then
        strStream = strStream & "|��ͳ��;" & mFeeBalance.��֧�� & ";0"
    End If
    
    סԺ�������_�Ͼ��� = strStream
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�Ͼ���(lng����ID As Long, lng����ID) As Boolean
    Dim strNO As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "select NO,���,��¼״̬,��¼���� from סԺ���ü�¼ where nvl(�Ƿ��ϴ�,0)=0 and ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ҽ�¼", lng����ID)
    Do Until rsTemp.EOF
        gstrSQL = "ZL_���˷��ü�¼_�ϴ�('" & rsTemp!NO & "'," & rsTemp!��� & "," & rsTemp!��¼���� & "," & rsTemp!��¼״̬ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select סԺ���� from ������Ϣ where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ҳid", lng����ID)
    
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ͼ��� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.������úϼ� & "," & mFeeBalance.������� + mFeeBalance.�����Ը� & ",0," & _
              mFeeBalance.ҽ����Χ���� & "," & mFeeBalance.ͳ��֧�� & "," & mFeeBalance.��֧�� & "," & _
              "0," & mFeeBalance.�����ʻ�֧�� & ",'" & mFeeBalance.סԺ��� & "'," & rsTemp!סԺ���� & ",null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���뱣���ʻ�")
    
    '����Ʊ��ʹ�ü�¼��סԺֻ������һ��Ʊ�ݣ�
'    zl_Ʊ��ʹ����ϸ_Insert (
'    ����ID_IN IN Ʊ��ʹ����ϸ.����ID%TYPE,
'    Ʊ��_IN IN Ʊ��ʹ����ϸ.Ʊ��%TYPE,
'    ����_IN IN Ʊ��ʹ����ϸ.����%TYPE,
'    ����_IN IN Ʊ��ʹ����ϸ.����%TYPE,
'    ԭ��_IN IN Ʊ��ʹ����ϸ.ԭ��%TYPE,
'    ����ID_IN IN Ʊ��ʹ����ϸ.����ID%TYPE,
'    ʹ��ʱ��_IN IN Ʊ��ʹ����ϸ.ʹ��ʱ��%TYPE,
'    ʹ����_IN IN Ʊ��ʹ����ϸ.ʹ����%TYPE
    If gblnBill Then
        strNO = GetNextBill(glng����ID)
        gstrSQL = "zl_Ʊ��ʹ����ϸ_Insert(" & glng����ID & ",3,'" & strNO & "',1,1," & lng����ID & ",sysdate,'" & UserInfo.���� & "')"
        gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
        Err.Raise 9000, gstrSysName, "����ҽ��ʹ��Ʊ�ݺţ�" & strNO
    End If
    
    סԺ����_�Ͼ��� = True
    Exit Function
errorhandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_�Ͼ���(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lng����ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����id", lng����ID)
    lng����ID = rsTemp!ID
    
    gstrSQL = "select * from ���ս����¼ where ��¼id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ԭʼ��¼", lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "���ս����¼��ԭʼ���ʵ��ݲ�����,�������˷�"
        Exit Function
    Else
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ͼ��� & "," & rsTemp!����ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!�������ý�� & "," & -rsTemp!ȫ�Ը���� & "," & -rsTemp!�����Ը���� & "," & -rsTemp!����ͳ���� & "," & -rsTemp!ͳ�ﱨ����� & "," & -rsTemp!���Ը���� & "," & _
              "0," & -rsTemp!�����ʻ�֧�� & ",'" & rsTemp!֧��˳��� & "'," & rsTemp!��ҳID & ",null,null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���ʼ�¼")
    End If
    
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        '���������ջط�Ʊ��¼
        gstrSQL = " Select * From Ʊ��ʹ����ϸ Where ����ID=" & lng����ID
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        With rsTemp
            Do While Not .EOF
                gstrSQL = "zl_Ʊ��ʹ����ϸ_Insert(" & !����ID & ",3,'" & !���� & "',2,2," & lng����ID & ",sysdate,'" & UserInfo.���� & "')"
                gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
                .MoveNext
            Loop
        End With
        gcnNJSYB.CommitTrans
        blnTrans = False
    End If
    סԺ�������_�Ͼ��� = True
    Exit Function
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Sub writeTxtFile(strFile As String, strWrite As String, Optional ByVal openFile As Boolean = True)
    Dim intSymbol As Long
    Dim strFolder As String
    
    On Error GoTo errorhandle
    Do Until InStr(intSymbol + 1, strFile, "\") = 0
        intSymbol = InStr(intSymbol + 1, strFile, "\")
        strFolder = Mid(strFile, 1, intSymbol)
        If Not mobjSystem.FolderExists(strFolder) Then mobjSystem.CreateFolder (strFolder)
    Loop

    If openFile Then                    '���ļ�
        If Not mobjSystem.FileExists(strFile) Then mobjSystem.CreateTextFile (strFile)
        Set mobjStream = mobjSystem.OpenTextFile(strFile, ForWriting)
        If strWrite <> "" Then          '��������ݽ���д��
            mobjStream.WriteLine (strWrite)
            mobjStream.Close
        End If
    Else
        If strWrite = "" Then
            mobjStream.Close
        Else
            mobjStream.WriteLine (strWrite)   '�����д�����ݵ��򿪱�־Ϊfalse,ֻ����д��
        End If
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mobjStream.Close
End Sub

Public Function readTxtFile(strFile As String) As String
    On Error GoTo errHandle
    
    If mobjSystem.FileExists(strFile) Then
        Set mobjStream = mobjSystem.OpenTextFile(strFile)
        readTxtFile = mobjStream.ReadLine
        mobjStream.Close
    End If
    Exit Function
    
errHandle:
    Err.Clear
    On Error Resume Next
    mobjStream.Close
End Function

Private Function fillSpa(strTemp As Variant, lngLen As Long, Optional fromRigth As Boolean = True) As String
    Dim lngStrLeng As Long
    Dim strStream As String
    Dim strUnion As String
    
    strTemp = IIf(IsNull(strTemp), "", Trim(strTemp))
    
    strUnion = StrConv(Trim(strTemp), vbFromUnicode)
    lngStrLeng = IIf(LenB(strUnion) > lngLen, lngLen, LenB(strUnion))
    strStream = IIf(LenB(strUnion) > lngLen, StrConv(LeftB(strUnion, 20), vbUnicode), strTemp)
    
    If fromRigth Then
        fillSpa = strStream & String(lngLen - lngStrLeng, " ")
    Else
        fillSpa = String(lngLen - lngStrLeng, " ") & strStream
    End If
End Function

Public Function analyseStr(strTemp As String, lngStart As Long, lngLen As Long) As String
    Dim strStream As String
    
    strStream = StrConv(UCase(strTemp), vbFromUnicode)
    
    analyseStr = Trim(StrConv(MidB(strStream, lngStart, lngLen), vbUnicode))
End Function

Public Function �������_�Ͼ���(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
'    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID='" & lng����ID & "' and ����=" & TYPE_�Ͼ���
'    Call OpenRecordset(rsTemp, gstrSysName)
'
'    If rsTemp.EOF Then
'        �������_�Ͼ��� = 100000
'    Else
'        �������_�Ͼ��� = IIf(rsTemp("�ʻ����") = 0, 100000, rsTemp("�ʻ����"))
'    End If
    �������_�Ͼ��� = 100000
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'���ܣ����û�����Ĳ��ݵ��ţ����ص���ĵ��š�
    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function FileExists(ByVal FileName As String, Optional ErrFlag As Boolean = True) As Boolean
    Dim Temp
    FileExists = True
    On Error Resume Next
proshow:
    Temp = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                If ErrFlag Then
                    If MsgBox("����û��׼���á�", vbInformation + vbRetryCancel, "����") = vbRetry Then
                        GoTo proshow:
                    End If
                End If
                FileExists = False
            End If
    End Select
End Function


Public Function �ҺŽ���_�Ͼ���(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'�����ߡ��������������÷���������ҺŲ�������
'����˵��������������ͨ������ҽ���̵�����ҺŽӿڣ��ֽⱾ�η�����ϸ���õ��������������ʻ����١�ҽ��������ٵȣ�������
'ע���������������������ڸ����ʻ���ҽ���������Ա��������Ҫ���ù���zl_���˽����¼_Update�Բ���Ԥ����¼������������
'���ù����嵥��˵����
'��������������
''*****************************************************************************
    �ҺŽ���_�Ͼ��� = True
End Function


Public Function �ҺŽ������_�Ͼ���(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'�����ߡ��������������÷���������ҺŲ�������
'����˵��������������ͨ������ҽ���̵�����Һų����ӿڣ��������ҺŽ��������
'���ù����嵥��˵����
'��������������
''*****************************************************************************
    �ҺŽ������_�Ͼ��� = True
End Function


Public Function ����ҽ����Ժ_�Ͼ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str˳��� As String) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim StrInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    
    ����ҽ����Ժ_�Ͼ��� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function AnalyServer(ByVal strConn As String) As String
    Dim arrData, arrColumn
    Dim intDO As Integer, intMAX As Integer
    
    strConn = UCase(Replace(strConn, """", ""))
    arrData = Split(strConn, ";")
    intMAX = UBound(arrData)
    For intDO = 0 To intMAX
        arrColumn = Split(arrData(intDO), "=")
        If arrColumn(0) = "SERVER" Then
            AnalyServer = arrColumn(1)
            Exit Function
        End If
    Next
End Function

Private Sub AnalyConf(strUser As String, strPass As String, strServer As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFileSys As New FileSystemObject
    Dim objStream As TextStream
    On Error GoTo errHand
    
    '�������ļ��ж�ȡҽ��ǰ�û����û�����������������
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Conf.ini"
    If objFileSys.FileExists(strFile) Then
        Set objStream = objFileSys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            Select Case arrLine(0)
            Case "USER"
                strUser = arrLine(1)
            Case "PASS"
                strPass = arrLine(1)
            Case "SERVER"
                strServer = arrLine(1)
            End Select
        Loop
        objStream.Close
        blnOpen = False
    End If
    
    If strUser = "" Then strUser = "zl9I_NJSYB"
    If strPass = "" Then strPass = "HIS"
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Private Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intnum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, Optional ByVal strBill As String) As Long
'���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
'������bytKind      =   Ʊ��
'      intNum       =   Ҫ��ӡ��Ʊ������
'      lngLastUseID =   �ϴ�ʹ�õ�����ID
'      lngShareUseID=   ���ز���ָ���Ĺ���ID
'      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
'���أ�
'      >0   =   �ɹ������õ�����ID
'      =0   =   ʧ��
'      -1   =   û������(����򲻹�����δ����),δ���ù���
'      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
'      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
'      -4   =   ָ�����ε�Ʊ�ݲ�����
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo ErrH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
                 "From Ʊ�����ü�¼ Where Ʊ��=" & bytKind & " And nvl(��ǰ����,��ʼ����)<>��ֹ���� And ID=" & lngLastUseID
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnNJSYB
        With rsTmp
            If .RecordCount > 0 Then    'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '����û�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intnum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
        
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSQL = "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = " & bytKind & " And nvl(��ǰ����,��ʼ����)<>��ֹ���� And ������ = '" & UserInfo.���� & "' And ʹ�÷�ʽ = 1" & vbNewLine & _
        "Order By ��ʼ����"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnNJSYB
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = "" & !ǰ׺�ı�
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
                 "From Ʊ�����ü�¼ Where Ʊ��=" & bytKind & " And nvl(��ǰ����,��ʼ����)<>��ֹ���� And ID=" & lngShareUseID
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnNJSYB
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    
    GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, Optional ByVal strBill As String) As Long
'���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
'������bytKind=Ʊ��
'      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
'      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
'˵����
'    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
'    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
'    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
'���أ�
'      ������Ʊ������ID>0
'      0=ʧ��
'      -1:û������(�����δ����)��Ҳû�й���(δ����)
'      -2:���õĹ���������
'      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As New ADODB.Recordset
    Dim rsSelf As New ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo ErrH
    
    '����Ա��ʣ�������Ʊ�ݼ�
    strSQL = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = " & bytKind & " And ʹ�÷�ʽ = 1 And nvl(��ǰ����,��ʼ����)<>��ֹ���� And ������ = '" & UserInfo.���� & "'" & vbNewLine & _
        "Order By ��ʼ����"
    If rsSelf.State = 1 Then rsSelf.Close
    rsSelf.CursorLocation = adUseClient
    rsSelf.Open strSQL, gcnNJSYB
    
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSQL = "Select ID,ʹ�÷�ʽ,ǰ׺�ı�,��ʼ����,��ֹ���� From Ʊ�����ü�¼ Where nvl(��ǰ����,��ʼ����)<>��ֹ���� And Ʊ��=" & bytKind & " And ID=" & lng����ID
        If rsSelf.State = 1 Then rsSelf.Close
        rsSelf.CursorLocation = adUseClient
        rsSelf.Open strSQL, gcnNJSYB
        If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                lngReturn = rsSelf!ID
            Else
                'û������ȡ����
                'If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '����Ʊ��
'            If rsTmp!ʣ������ > 0 Then
                '��ʣ��
                lngReturn = rsTmp!ID
'            Else
'                '������ʣ�������
'                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
'                lngReturn = rsSelf!ID
'            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Private Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
'      2.�ſ��ѱ���ĺ���
    Dim rsMain As New ADODB.Recordset
    Dim rsDelete As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo ErrH
    
    strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where nvl(��ǰ����,��ʼ����)<>��ֹ���� And ID=" & lng����ID
    If rsMain.State = 1 Then rsMain.Close
    rsMain.CursorLocation = adUseClient
    rsMain.Open strSQL, gcnNJSYB
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(IncStr(rsMain!��ǰ����))
    End If
    
    strSQL = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����=1 And ԭ��=5 And ����>='" & strBill & "' And ����ID=" & lng����ID & _
        " Order by ����"
    If rsDelete.State = 1 Then rsDelete.Close
    rsDelete.CursorLocation = adUseClient
    rsDelete.Open strSQL, gcnNJSYB
    Do While True
        '��鷶Χ
        If Left(strBill, Len("" & rsMain!ǰ׺�ı�)) <> UCase("" & rsMain!ǰ׺�ı�) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!��ʼ����) And strBill <= UCase(rsMain!��ֹ����)) Then
            Exit Function
        End If
                
        '�ſ������
        rsDelete.Filter = "����='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Function AnalyBill(ByVal rsDetail As ADODB.Recordset) As Long
    Dim intBills_A As Integer, intBills_B As Integer
    Dim int��ϸ�� As Integer, int�վݷ�Ŀ As Integer
    Dim str�վݷ�Ŀ As String
    Dim rsƱ�ݷ�Ŀ As New ADODB.Recordset
    '���ر��ν����Ʊ������
    
    With rsDetail
        int��ϸ�� = .RecordCount
        Do While Not .EOF
            On Error Resume Next
            Err = 0
            'ȡ��ӡʱӦ��ʹ�õ��վݷ�Ŀ��V10��18���������ű�
            gstrSQL = "select �վݷ�Ŀ from �վݷ�Ŀ��Ӧ a,�շѼ�Ŀ b where a.����=0 and a.������Ŀid=b.������Ŀid and (b.��ֹ����>sysdate or b.��ֹ����=to_date('3000-01-01','yyyy-mm-dd') )" & _
                     " and ִ������<sysdate and b.�շ�ϸĿid= [1]"
            Set rsƱ�ݷ�Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, "Ʊ�ݷ�Ŀ", CLng(!�շ�ϸĿID))
            If Err = 0 Then
                If InStr(1, str�վݷ�Ŀ, rsƱ�ݷ�Ŀ!�վݷ�Ŀ) = 0 Then
                    str�վݷ�Ŀ = str�վݷ�Ŀ & "," & rsƱ�ݷ�Ŀ!�վݷ�Ŀ
                    int�վݷ�Ŀ = int�վݷ�Ŀ + 1
                End If
            Else
                If InStr(1, str�վݷ�Ŀ, !�վݷ�Ŀ) = 0 Then
                    str�վݷ�Ŀ = str�վݷ�Ŀ & "," & !�վݷ�Ŀ
                    int�վݷ�Ŀ = int�վݷ�Ŀ + 1
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '���㱾��Ʊ��ʹ������,ȡ���
    intBills_A = int��ϸ�� \ gint��ϸ��
    If int��ϸ�� Mod gint��ϸ�� <> 0 Then intBills_A = intBills_A + 1
    intBills_B = int�վݷ�Ŀ \ gint�վݷ�Ŀ
    If int�վݷ�Ŀ Mod gint�վݷ�Ŀ <> 0 Then intBills_B = intBills_B + 1
    
    If intBills_A >= intBills_B Then
        AnalyBill = intBills_A
    Else
        AnalyBill = intBills_B
    End If
End Function

'���ʣ�µ�Ʊ�������Ƿ��ã�����������ʾ����������н���
Private Function IsEnough() As Boolean
    Dim lng��ǰ���� As Long, lng��ֹ���� As Long
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " Select ǰ׺�ı�,��ֹ���� From Ʊ�����ü�¼ Where ID=" & glng����ID
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnNJSYB
    lng��ǰ���� = Mid(GetNextBill(glng����ID), Len(rsTemp!ǰ׺�ı�) + 1)
    lng��ֹ���� = Mid(rsTemp!��ֹ����, Len(rsTemp!ǰ׺�ı�) + 1)
    IsEnough = (lng��ֹ���� - lng��ǰ���� + 1 >= gintBills)
End Function

Private Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function
