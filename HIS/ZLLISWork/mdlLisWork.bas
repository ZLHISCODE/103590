Attribute VB_Name = "mdlLisWork"
Option Explicit

Public gobjEmrInterface As Object           '�°没�����븽���ȡ����
Public gobjpublicExpenses As Object

Public Enum COLOR
    
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ��ɫ = &H40C0&
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
End Enum
'    ��������ɫ = &HFF&
'    ����ǰ��ɫ = &H8000000F
'    ���걳��ɫ = &H80FF&
'    ����ǰ��ɫ = &H80000008
'    Ĭ��ǰ��ɫ = &H80000008


'ȡ�õ�ǰ�������е�SQL��select Wmsys.Wm_Concat(Column_Name) From User_Tab_Columns Where Table_Name = Upper('�������')
Public Const gConst_������Ϣ_���� As String = "a.����id,a.�����,a.סԺ��,a.���￨��,a.����֤��,a.�ѱ�,a.ҽ�Ƹ��ʽ,a.����,a.�Ա�,a.����,a.��������," & _
                                              "a.�����ص�,a.���֤��,a.���,a.ְҵ,a.����,a.����,a.����,a.ѧ��,a.����״��,a.��ͥ��ַ,a.��ͥ�绰,a.��ͥ��ַ�ʱ�," & _
                                              "a.��ϵ�˹�ϵ,a.��ϵ�˵�ַ,a.��ϵ�˵绰,a.��ͬ��λID,a.������λ,a.��λ�绰,a.��λ�ʱ�,a.��λ������,a.��λ�ʺ�," & _
                                              "a.������,a.��������,a.����ʱ��,a.����״̬,a.��������,a.סԺ����,a.��ǰ����ID,a.��ǰ����ID,a.��Ժʱ��,a.��Ժʱ��," & _
                                              "a.IC����,a.������,a.����,a.�Ǽ�ʱ��,a.ͣ��ʱ��,a.��ǰ����,a.ҽ����,a.��ѯ����,a.��Ժ,a.����֤��,a.�໤��,a.����,a.��ҳid"

                                              
Public Const gConst_�������_���� As String = " a.����ID,a.����,a.Ԥ�����,a.������� "

Public Const gConst_��������_���� As String = "a.����,a.���Ƶ��,a.���ʱ��,a.���巽ʽ,a.�հ���ʽ,a.�����ʿ�ͼ,a.����ʱָ������,a.�ʿ�ˮƽ��,a.�ϴ��ʿ���,a.QC��," & _
                                              "a. �Լ���Դ,a.У׼����Դ,a.ID,a.����,a.����,a.����,a.���Ӽ����,a.ͨѶ������,a.ͨѶ�˿�,a.������,a.����,a.ֹͣλ," & _
                                              "a.У��λ,a. ��������,a.������־ɫ,a.ʹ��С��ID,a.�ʿر걾��,a.��ע,a.΢����,a.ת������,a.ת������ID,a.�ʿ�����,a.���ڵ�λ"
                                              
Public Const gConst_����ҽ������_���� As String = "a.��������,a.���ʱ��,a.�ͼ���,a.�����ӡ,a.�زɱ걾,a.�걾�ͳ�ʱ��,a.�����,a.ҽ��ID,a.���ͺ�,a.��¼����,a.NO," & _
                                                  "a. ��¼���,a.��������,a.������,a.����ʱ��,a.�״�ʱ��,a.ĩ��ʱ��,a.ִ��״̬,a.ִ�в���ID,a.�Ʒ�״̬,a.ִ�м�," & _
                                                  "a.ִ�й���,a. ������,a.����ʱ��,a.��������,a.����ID,a.�������,a.����ʱ��,a.ִ��˵��,a.������,a.����ʱ��,a.����ʱ��"
                                                  
Public Const gConst_����걾��¼_���� As String = "a.��ҳID,a.������Ŀ,a.��������,a.������,a.����ʱ��,a.��ʶ��,a.����,a.���˿���,a.����," & _
                                                  "a.������,a. ����ʱ��,a.һ������,a.��������,a.��������,a.���δͨ��,a.��������,a.���䵥λ,a.����," & _
                                                  "a.�Һŵ�,a.�����,a. סԺ��,a.��������,a.ID,a.ҽ��ID,a.�걾���,a.����ʱ��,a.������,a.�걾����,a.������," & _
                                                  "a.����ʱ��,a.����״̬,a.������,a. ����ʱ��,a.�����,a.���ʱ��,a.�ϲ������,a.��ӡ����,a.��������,a.����ID," & _
                                                  "a.��������,a.������,a.��ע,a.δͨ�����ԭ��,a. ����ʱ��,a.�걾��̬,a.�Ƿ��ʿ�Ʒ,a.ִ�п���ID,a.΢����걾," & _
                                                  "a.NO,a.�Ƿ���,a.�걾���,a.���鱸ע,a.������,a.�������ID,a. ������Դ,a.����ID,a.Ӥ��,a.����,a.�Ա�,a.����,a.�ϲ�ID"
                                                  
Public Const gconst_����ҽ����¼_���� As String = "a.�ɷ����,a.���δ�ӡ,a.��鷽��,a.ִ�б��,a.�ͼ���,a.�ϴδ�ӡʱ��,a.����,a.��˱��,a.��Ҫ,a.ִ������," & _
                                                  "a.������־,a.��ʼִ��ʱ��,a.ִ����ֹʱ��,a.�ϴ�ִ��ʱ��,a.��������ID,a.����ҽ��,a.����ʱ��,a.У�Ի�ʿ,a.У��ʱ��," & _
                                                  "a.ͣ��ҽ��,a.ͣ��ʱ��,a.ȷ��ͣ��ʱ��,a.����ID,a.�Ƿ��ϴ�,a.�����,a.�״α���ʱ��,a.ժҪ,a.ID,a.���ID,a.ǰ��ID," & _
                                                  "a.������Դ,a.����ID,a.��ҳID,a.�Һŵ�,a.Ӥ��,a.����,a.�Ա�,a.����,a.���˿���ID,a.���,a.ҽ��״̬,a.ҽ����Ч," & _
                                                  "a.�������,a.������ĿID,a.�걾��λ,a.�շ�ϸĿID,a.����,a.��������,a.�ܸ�����,a.ҽ������,a.ҽ������,a.ִ�п���ID," & _
                                                  "a.Ƥ�Խ��,a.ִ��Ƶ��,a.Ƶ�ʴ���,a.Ƶ�ʼ��,a.�����λ,a.ִ��ʱ�䷽��,a.�Ƽ�����"


Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrProductName As String            '��Ʒ��ƣ����磺����
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gintSelectFocus As Integer           '���ڿؼ����������⣬��Ϊѡ����ȷ����Ľ���
                                            '1=Dkp_ID_List;2=frmLabRequest;3=frmLisStationWrite;4=frmLisStationWrite2(1)
                                            '5=frmLisStationWrite2(2)

'ҽ������
Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '�Ƿ�����ҽ��
Public gintInsure As Integer

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjRichEPR As New cRichEPR          '�������Ĳ���
Public gbytCardNOLen As Long                '���￨����

Public gobjEmr As Object                    '��������


Public gstrUnitName As String               '�û���λ����
Public gfrmMain As Object

Public gstrSql As String
Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gblnOK As Boolean
Public gLabcboDept As Object

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO


'HISϵͳ����

Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"


Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�����¼���� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
End Enum

'����ǩ��
Public gintCA As Integer '����ǩ����֤����
Public gstrESign As String '����ǩ�����Ƴ���
Public gobjESign As Object '����ǩ���ӿڲ���

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const LONG_MAX = 2147483647 'Long�����ֵ
Public Const CuvetteNumberLen = 12 '�Թ����볤��

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public grsSysPars As ADODB.Recordset
Public gbln������Ȩ���� As Boolean  '�Ƿ�����������ҽʦ��Ȩ����


Public gblnManualPH As Boolean '�ֹ�ʹ��������Ϊ�걾��
Public gintNumberPH As Integer 'ÿ�����걾��
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip
Private mlngImageID As Long  '�������¶�ȡ��ͬͼƬ
Public gblnִ�к���� As Boolean    'ִ�к��Զ���˻��۵�
Public mobjLisInsideComm As Object                                      'LIS�ڲ��ӿ�
Public mobjZLIHISPlugIn As Object                                       'ZLHISͨ�ò���ӿ�
Private mintWarn As Integer                                             '-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ַ����ļ���")
    With rsTmp
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function


Public Sub NewColumn(msf As Object, ByVal vText As String, Optional ByVal vWidth As Single = 1200, Optional ByVal vAlignment As Byte = 9)
    Dim i As Long
    
    msf.Cols = msf.Cols + 1
    i = msf.Cols - 1
    
    msf.TextMatrix(0, i) = vText
    msf.ColWidth(i) = vWidth
    msf.ColAlignment(i) = vAlignment
    
    On Error Resume Next
    msf.ColAlignmentFixed(i) = vAlignment
    
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    On Error GoTo errH
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    X = objPoint.X * 15 + objBill.CellLeft
    Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal blnMerge As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '   blnMerge���Ƿ�ϲ���ͬID���С�����ϲ�����������ͬ����ֵ�ԡ�;���ָ�
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngloop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngloop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        lngCurrRow = -1
        If blnMerge Then lngCurrRow = FindGridLine(objMsf, CStr(zlCommFun.Nvl(rsData("ID"))))
        If lngCurrRow = -1 Then
            lngRow = lngRow + 1
            lngCurrRow = lngRow
        End If
        
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
        
        On Error GoTo ErrHand
        
        For lngloop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngloop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngloop)
                                        
                On Error GoTo ErrHand
                
                strOldValue = objMsf.TextMatrix(lngCurrRow, lngloop)
                If strMask <> "" Then
                    strNewValue = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop))), strMask)
                Else
                    strNewValue = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop)))
                End If
                objMsf.TextMatrix(lngCurrRow, lngloop) = IIf(Trim(strOldValue) = "", strNewValue, _
                     strOldValue & IIf(InStr(";" & strOldValue & ";", ";" & strNewValue & ";") > 0, "", ";" & strNewValue))
            End If
            
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function FindGridLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '����:����RowData����strSeekID����
    '����:
    '����:�кŻ�-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    FindGridLine = -1
    For i = 1 To objMsf.Rows - 1
        If objMsf.RowData(i) = strSeekID Then Exit For
    Next
    If i <= objMsf.Rows - 1 Then FindGridLine = i
End Function

Public Function FillListData(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngloop As Long
    
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rs("ǰ��ɫ").Name = "ǰ��ɫ")
    
    On Error GoTo ErrHand
    
    LockWindowUpdate objLvw.hWnd
    
    
    Do While Not rs.EOF
        
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("����").Value, rs("ͼ��").Value, rs("ͼ��").Value)
        For lngloop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngloop - 1) = zlCommFun.Nvl(rs(objLvw.ColumnHeaders(lngloop).Text).Value)
        Next
        
        If blnForeColor Then
            objItem.ForeColor = Val(rs("ǰ��ɫ").Value)
            For lngloop = 2 To objLvw.ColumnHeaders.Count
                objItem.ListSubItems(lngloop - 1).ForeColor = objItem.ForeColor
            Next
        End If
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillListData = True
    
    Exit Function
ErrHand:
    LockWindowUpdate 0
    If ErrCenter = 1 Then Resume
End Function


Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789<>", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.-<>+Ee", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    Dim lngloop As Long
    
    Select Case bytMode
    Case 1
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 99
        For lngloop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngloop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    Set rs = zlDatabase.OpenSQLRecord("SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlCISBase")
    GetMaxLength = rs.Fields(0).DefinedSize
    
End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'����: װ��������ָ�������������������е���������
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub


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

Public Sub ResetVsf(objVsf As Object)
    objVsf.Rows = 1
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function CheckIsAllowAuditing(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ������Ƿ��������,���Ƿ�������˵�����
    '--------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
        
    strSQL = "SELECT ROWNUM AS ���,A.����ʱ�� AS �걾ʱ��,A.�걾���, D.������ AS ������Ŀ, A.������ AS ���ν��,B.������ AS �ϴν�� " & _
                 "FROM (SELECT A.����ʱ��,A.�걾���, B.������Ŀid, B.������, A.����ʱ�� " & _
                         "FROM ����걾��¼ A, ������ͨ��� B " & _
                        "WHERE A.ID = B.����걾ID AND A.������ = B.��¼���� AND A.ID = [1]) A, " & _
                      "(SELECT C.������Ŀid, C.������, A.����ʱ�� " & _
                         "FROM ����걾��¼ A,����ҽ����¼ B,������ͨ��� C, " & _
                              "(SELECT B.����ID, B.��ҳID " & _
                                 "FROM ����걾��¼ A, ����ҽ����¼ B " & _
                                "WHERE A.ҽ��ID + 0 = B.ID AND A.ID = [1] ) D " & _
                        "WHERE (C.������Ŀid,A.����ʱ��) IN (SELECT D.������Ŀid,MAX(A.����ʱ��) " & _
                                           "FROM ����걾��¼ A,����ҽ����¼ B,������ͨ��� D," & _
                                                "(SELECT B.����ID, B.��ҳID, A.����ʱ�� " & _
                                                   "FROM ����걾��¼ A, ����ҽ����¼ B,������ͨ��� C " & _
                                                  "WHERE C.����걾ID=A.ID AND A.������=C.��¼���� AND A.ҽ��ID + 0 = B.ID AND A.ID = [1] ) C " & _
                                          "WHERE A.����ʱ�� < C.����ʱ�� AND A.ҽ��ID = B.ID AND D.������Ŀid=C.������Ŀid AND A.������=D.��¼���� AND D.����걾ID=A.ID AND " & _
                                                "B.����ID = C.����ID AND NVL(B.��ҳID,0) = NVL(C.��ҳID,0) GROUP BY D.������Ŀid) AND " & _
                              "A.ҽ��ID = B.ID AND B.����ID = D.����ID AND " & _
                              "NVL(B.��ҳID,0) = NVL(D.��ҳID,0) AND C.����걾ID = A.ID AND " & _
                              "A.������ = C.��¼����) B, " & _
                      "������Ŀ C,����������Ŀ D " & _
                "WHERE A.������Ŀid = B.������Ŀid(+) AND C.������� = 1 AND " & _
                      "C.����쳣���� IS NOT NULL AND C.������Ŀid = D.ID AND " & _
                      "C.������Ŀid = A.������Ŀid AND " & _
                      "(A.����ʱ�� - B.����ʱ��) <=TO_NUMBER(SUBSTR(C.����쳣����, 1, INSTR(C.����쳣����, ';') - 1)) AND " & _
                      "ABS(TO_NUMBER(A.������) - TO_NUMBER(B.������)) >=TO_NUMBER(SUBSTR(C.����쳣����,INSTR(C.����쳣����, ';') + 1,LENGTH(C.����쳣����) - INSTR(C.����쳣����, ';')))"
                      
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngKey)
    
    CheckIsAllowAuditing = (rs.BOF = True)
    If rs.BOF = False Then
        CheckIsAllowAuditing = frmLisStationError.ShowError(frmMain, rs)
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��
    '����:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = zlDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ָ����ʼ����"
        If bytFlag = 1 Then
            GetDateTime = zlDatabase.GetPara("���μ��鷶Χָ����ʼ����", 100, 1208, Format(dateNow - 30, "yyyy-mm-dd 00:00:00"))
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "���ظ�"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    Case "�Զ���"
        GetDateTime = "�Զ���"
    End Select
    
End Function

Public Sub ApplyResultColor(vsf As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    Dim lngReferenceLow As Long                             '�ο�����ɫ
    Dim lngReferenceHigh As Long                            '�ο�����ɫ
    Dim lngReferenceExigency As Long                        '�ο���ʾ��ɫ
    
    '��ȡ��ɫ
    lngReferenceLow = Val(zlDatabase.GetPara("�ο���ɫ_ƫ��", 100, 1208, 0))
    If lngReferenceLow = 0 Then lngReferenceLow = 8454143
    lngReferenceHigh = Val(zlDatabase.GetPara("�ο���ɫ_ƫ��", 100, 1208, 0))
    If lngReferenceHigh = 0 Then lngReferenceHigh = 8438015
    lngReferenceExigency = Val(zlDatabase.GetPara("�ο���ɫ_��ʾ", 100, 1208, 0))
    If lngReferenceExigency = 0 Then lngReferenceExigency = 16576
    
    Select Case bytMode
        Case 0, 1
            lngColor = &H80000005
            lngForeColor = COLOR.Ĭ��ǰ��ɫ
        Case 5, 6 '�쳣�͡���
            lngColor = lngReferenceExigency
            lngForeColor = COLOR.����ǰ��ɫ
        Case 2
            lngColor = lngReferenceLow
            lngForeColor = COLOR.����ǰ��ɫ
        Case Else
            lngColor = lngReferenceHigh
            lngForeColor = COLOR.����ǰ��ɫ
    End Select
    
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
    
    
End Sub

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Sub ClearGrid(vsf As Object, Optional ByVal Row As Long = 1)
    '--------------------------------------------------------------------------------------------------------
    '����:����������
    '--------------------------------------------------------------------------------------------------------
    vsf.Rows = Row + 1
    vsf.RowData(Row) = 0
    vsf.Cell(flexcpText, Row, 0, Row, vsf.Cols - 1) = ""
    
End Sub

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub DeleteRecord(rs As ADODB.Recordset)
    '-----------------------------------------------------------------------------------
    '����:ɾ����¼��
    '����:rs        Ҫɾ���ļ�¼��
    '����:��
    '-----------------------------------------------------------------------------------
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
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

Public Function GetReportCode(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����;
    '--------------------------------------------------------------------------------------------------------
    Dim rsPaitentType As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lngҽ��ID = 0 And lng���ͺ� = 0 Then Exit Function
    
'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                       "A.NO," & _
                       "A.��¼���� " & _
                "FROM ����ҽ������ A,�����ļ��б� C,����ҽ����¼ D,��������Ӧ�� E " & _
                "Where E.�����ļ�id = C.ID " & _
                        "AND D.������ĿID=E.������ĿID " & _
                      "AND A.ҽ��ID=D.ID AND E.Ӧ�ó���=Decode(D.������Դ,2,2,4,4,1) " & _
                      " AND D.���id= [1] "
    strSQL = "select b.�������� from ����ҽ����¼ A,������ҳ B  where a.����id=b.����id and a. ���id = [1]"
    Set rsPaitentType = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngҽ��ID)
    
    If rsPaitentType.RecordCount > 0 Then
        strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������, A.NO, A.��¼����, F.ID, F.����" & vbNewLine & _
                "From ����ҽ������ A, �����ļ��б� C, ����ҽ����¼ D, ��������Ӧ�� E, ������ĿĿ¼ F,������ҳ G" & vbNewLine & _
                "Where E.�����ļ�id = C.ID And D.������Ŀid = E.������Ŀid And D.������Ŀid = F.ID And A.ҽ��id = D.ID and d.����id=g.����id And" & vbNewLine & _
                "      E.Ӧ�ó��� = Decode(D.������Դ, 2, Decode(g.��������,1,1,2), 4, 4, 1) And D.���id = [1] " & vbNewLine & _
                "Order By F.���� "
    Else
        
        strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������, A.NO, A.��¼����, F.ID, F.����" & vbNewLine & _
                "From ����ҽ������ A, �����ļ��б� C, ����ҽ����¼ D, ��������Ӧ�� E, ������ĿĿ¼ F" & vbNewLine & _
                "Where E.�����ļ�id = C.ID And D.������Ŀid = E.������Ŀid And D.������Ŀid = F.ID And A.ҽ��id = D.ID And" & vbNewLine & _
                "      E.Ӧ�ó��� = Decode(D.������Դ, 2, 2, 4, 4, 1) And D.���id = [1] " & vbNewLine & _
                "Order By F.���� "
    End If
                          
    If DataMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If

'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
'                       "A.NO," & _
'                       "A.��¼���� " & _
'                "FROM ��������Ӧ�� A,�����ļ�Ŀ¼ C,����ҽ����¼ D,����ҽ������ B " & _
'                "Where A.�����ļ�id = C.ID " & _
'                      "AND A.������Ŀid=D.������ĿID " & _
'                      "AND B.����ID=D.����ID " & _
'                      "AND NVL(B.��ҳID,0)=NVL(D.��ҳID,0) " & _
'                      "AND B.�ļ�id=C.ID " & _
'                      "AND D.���id=" & lngҽ��id & " " & _
'                      "AND A.���ͺ�=" & lng���ͺ�

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngҽ��ID, lng���ͺ�)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.Nvl(rs("������"))
        strNO = zlCommFun.Nvl(rs("NO"))
        bytMode = zlCommFun.Nvl(rs("��¼����"), 1)
    End If
    
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As Boolean
    '�����շ�״̬
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strSQLbak As String
    Dim intPatientType As Integer               '������Դ
    On Error GoTo errH
    
    CheckChargeState = False
    
    strSQL = "select ������Դ from ����걾��¼ where id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "��������", lngKey)
    If rs.EOF = True Then Exit Function
    intPatientType = rs("������Դ")
    
    If blnOrder Then
        strSQL = _
            "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                  "from סԺ���ü�¼ A, " & _
                  "( " & _
                       "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id))  " & _
                       "Union " & _
                       "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id)) " & _
                  ") B " & _
                "Where A.NO = B.NO "
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
        End If
    Else
        strSQL = _
            "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                  "from סԺ���ü�¼ A, " & _
                  "( " & _
                       "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                       "Union " & _
                       "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                  ") B " & _
                "Where A.NO = B.NO and mod(a.��¼����,10) = b.��¼���� "
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
        End If
    End If
    
    strSQL = strSQL & " Order by ��¼״̬ "
    If DataMoved Then
        strSQL = Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����걾��¼", "H����걾��¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

    If rs.BOF Then Exit Function
    If rs("��¼״̬").Value = 0 Then Exit Function
    
    CheckChargeState = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadAdvicePrice(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal str�ѱ� As String, Optional ByVal DataMoved As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ��ҽ���ļƼ۹�ϵ����ʱ��¼��
    '˵����Ҫ�������ĿӦ�ò��Ƕ���,Ժ��ִ��,����Ʒ�
    '----------------------------------------------------------------------------------------------------
    Dim rsAdvice As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
        
    On Error GoTo errH
            
    '��ȡҪ���������õ�ҽ����¼(������������,��鲿λ������������)
    strSQL = _
        " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID," & _
        " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����" & _
        " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
        " Where B.���ID= [1] " & _
        " And A.ҽ��ID=B.ID And A.���ͺ�+0=" & lng���ͺ� & _
        " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
        " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,A.��������"
    strSQL = strSQL & " Union ALL " & _
        " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID," & _
        " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����" & _
        " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
        " Where B.ID= [1] " & _
        " And A.ҽ��ID=B.ID And A.���ͺ�+0=" & lng���ͺ� & _
        " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
        " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,A.��������" & _
        " Order by ���"
    If DataMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
            
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngҽ��ID)
    
    For i = 1 To rsAdvice.RecordCount

        strSQL = _
            " Select 1 " & _
            " From �����շѹ�ϵ A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,������Ŀ D" & _
            " Where A.������ĿID= [1] " & _
            " And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And B.������ĿID=D.ID" & _
            " And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
            " And Nvl(A.���ж���,0)=1 And Nvl(C.�Ƿ���,0)=0"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", Val(rsAdvice!������ĿID))
        
        If rsTmp.RecordCount > 0 Then
            LoadAdvicePrice = True
            Exit Function
        End If
        
        rsAdvice.MoveNext
    Next
        
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetColNumber(objVsf As Object, ByVal strCaption As String) As Long
    
    Dim lngloop As Long
    
    GetColNumber = -1
    
    For lngloop = 0 To objVsf.Cols - 1
        If objVsf.TextMatrix(0, lngloop) = strCaption Then
            GetColNumber = lngloop
            Exit Function
        End If
    Next
    
End Function

Public Sub VsfCellFormat(objVsf As Object, ByVal lngCol As Long, ByVal strFormat As String, Optional ByVal iType As Integer = -1, Optional ByVal iTypeCol As Integer = -1)
    'iType�����ʽ������������
    '  0�����֡�1���ַ���2�����ڡ�3���߼���-1�����ޣ�ȱʡ��
    'iTypeCol���������͵Ĵ洢�ֶ����
    Dim lngloop As Long
    On Error GoTo errH
    For lngloop = 1 To objVsf.Rows - 1
        If iType = -1 Then
            objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
        Else
            If iTypeCol = -1 Then
                If iType = 0 And IsNumeric("-" & objVsf.TextMatrix(lngloop, lngCol)) Then objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
            Else
                If iType = Val(objVsf.TextMatrix(lngloop, iTypeCol)) Then objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '��(X1,Y1),(X2,Y2)֮��ʹ��ForeColorɫ��һֱ��
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    On Error GoTo errH
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngloop = 0 To UBound(varArray)
        varItem = Split(varArray(lngloop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngloop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errH
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngloop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngloop) = False Then
            lngLastRow = lngloop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.�������е���
    For lngloop = 1 To objLineX.UBound
        objLineX(lngloop).Visible = False
    Next
    
    For lngloop = 1 To objLineY.UBound
        objLineY(lngloop).Visible = False
    Next
    
    '2.���¼�����Ҫ������
    For lngloop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngloop Then Load objLineY(lngloop)

        With objLineY(lngloop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngloop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function



Public Function ShowGrdFilterDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;��ʾ�ı�����ѡ���б�(ֻ���ڱ��ؼ�)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo ErrHand

    If InStr(objVsf.EditText, "'") > 0 Then Exit Function
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
            
    'ִ�в�ѯ
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If rs.BOF Then
        If blnPrompt Then MsgBox "û���ҵ���ƥ��Ľ����", , gstrSysName
        Exit Function                            'û�н����ֱ�ӷ���
    End If
            
    If rs.RecordCount = 1 And blnFilter Then GoTo Over                    '��Ϊ��������ң����ֻ��һ������ֱ�ӷ���
    If frmSelectList.ShowSelect(frmParent, rs, strLvw, lngX, lngY, lngCX, lngCY, strSavePath, strDescrible, , , objVsf.CellHeight) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rs
    
    ShowGrdFilterDialog = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowGrdSelectDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������+�б�ṹ,Ӧ���ڱ��ؼ�
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo ErrHand
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If rs.BOF Then
        MsgBox "û�п�ѡ������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objVsf.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
    
    
    If frmSelectExplorer.ShowSelect(frmParent, rs, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, strSavePath, strLvw, strDescrible) Then
                        
        Set rsResult = rs
        ShowGrdSelectDialog = True
        
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Sub LocationVsf(objVsf As Object, ByVal lngRow As Long, ByVal lngCol As Long)
    
    On Error Resume Next
    
    objVsf.Row = lngRow
    objVsf.Col = lngCol
    objVsf.ShowCell objVsf.Row, objVsf.Col
    objVsf.SetFocus
End Sub

Public Function CheckNumeric(ByVal strText As String, ByVal lngLength As Long, Optional ByVal lngDecLength As Long = 0, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '����:����ַ�������ֵ��Ч��
    '--------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    
    Dim str�������� As String
    Dim strС������ As String
    On Error GoTo errH
    If lngDecLength = 0 Then
        '����
        Select Case bytMode
        Case 1      '������
            str�������� = strText
        Case 2      '������
            If Left(strText, 1) <> "-" And strText <> "0" Then
                CheckNumeric = "ӦΪ���������㣡"
                Exit Function
            End If
            str�������� = Mid(strText, 2)
            
        Case 3      '��������
            If Left(strText, 1) = "-" Then str�������� = Mid(strText, 2)
        End Select
    Else
        'С��
        Select Case bytMode
        Case 1      '��С��
            If Len(strText) > lngLength + 1 Then
                CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '��С������
                str�������� = Left(strText, InStr(strText, ".") - 1)
                strС������ = Mid(strText, InStr(strText, ".") + 1)
            Else
                str�������� = strText
            End If
            
        Case 2      '��С��
            If Len(strText) > lngLength + 2 Then
                CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                Exit Function
            End If
            
            If Left(strText, 1) <> "-" Then
                CheckNumeric = "���Ǹ�����"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '��С������
                str�������� = Mid(strText, 2, InStr(strText, ".") - 2)
                strС������ = Mid(strText, InStr(strText, ".") + 1)
            Else
                str�������� = Mid(strText, 2)
            End If
            
        Case 3      '����С��
            If Left(strText, 1) = "-" Then
                If Len(strText) > lngLength + 2 Then
                    CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '��С������
                    str�������� = Mid(strText, 2, InStr(strText, ".") - 2)
                    strС������ = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str�������� = Mid(strText, 2)
                End If
            Else
                If Len(strText) > lngLength + 1 Then
                    CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '��С������
                    str�������� = Mid(strText, 1, InStr(strText, ".") - 1)
                    strС������ = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str�������� = strText
                End If
                
            End If
        End Select
    End If
    
    If Len(str��������) > (lngLength - lngDecLength) Then
        If lngDecLength = 0 Then
            CheckNumeric = "���ȳ�����" & (lngLength - lngDecLength) & "λ��"
        Else
            CheckNumeric = "�������ݳ��ȳ�����" & (lngLength - lngDecLength) & "λ��"
        End If
        Exit Function
    End If
    
    If Len(strС������) > lngDecLength Then
        CheckNumeric = "С�����ݳ��ȳ�����" & lngDecLength & "λ��"
        Exit Function
    End If
    
    For lngloop = 1 To Len(str��������)
        If Mid(str��������, lngloop, 1) < "0" Or Mid(str��������, lngloop, 1) > "9" Then
            CheckNumeric = "ӦΪ�����ͣ�"
            Exit Function
        End If
    Next
    
    For lngloop = 1 To Len(strС������)
        If Mid(strС������, lngloop, 1) < "0" Or Mid(strС������, lngloop, 1) > "9" Then
            CheckNumeric = "ӦΪ�����ͣ�"
            Exit Function
        End If
    Next
    
    
    CheckNumeric = ""
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �ϴ����� From zlDataMove Where ϵͳ=[1] And ���=1 And �ϴ����� is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '�ϴ�����û��ʱ��,"<"�ж���ת��������һ��
        If vDate < rsTmp!�ϴ����� Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetConnectDevs() As String
'���ܣ���ȡ�������ӵļ�����������;�ָ�
'��ͮ��
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long
    
    GetConnectDevs = ""
    On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
    If Not IsEmpty(aPorts) Then
        For i = 0 To UBound(aPorts)
            PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
            lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
            If lngDeviceID > 0 Then
                GetConnectDevs = GetConnectDevs & ";" & lngDeviceID
            End If
        Next
        If Len(GetConnectDevs) > 0 Then GetConnectDevs = Mid(GetConnectDevs, 2)
    End If
End Function

Public Function FindComboItem(objCombox As Object, ByVal lngFind As Long) As Integer
    Dim i As Integer
    
    For i = 0 To objCombox.ListCount - 1
        If objCombox.ItemData(i) = lngFind Then Exit For
    Next
    If i > objCombox.ListCount - 1 Then i = -1
    
    FindComboItem = i
End Function

'---����Ϊֱ���������

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    On Error GoTo errH
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    On Error GoTo errH
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    If Msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
    End If
End Function

Public Function CheckOneDuty(ByVal strҽ�� As String, ByVal strְ�� As String, ByVal strҽ�� As String, ByVal blnҽ�� As Boolean) As String
'���ܣ���鵱ǰָ��ҩƷ����ְ���Ƿ����
'������strҽ��=ҩƷҽ����ʾ����
'      strְ��=ҩƷ����ְ��
'      strҽ��=����ҽ��
'      blnҽ��=�Ƿ񹫷ѻ�ҽ������
'      grsDuty=��¼ҽ��ְ�񻺴�
'���أ�ְ���������ʾ��Ϣ����������򷵻ؿա�
    Const STR_ְ�� = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim intְ��A As Integer, intְ��B As Integer
    
    If Len(strְ��) <> 2 Or strҽ�� = "" Then Exit Function
    
    'ȡҩƷ����ְ��
    If blnҽ�� Then
        intְ��B = Val(Right(strְ��, 1))
    Else
        intְ��B = Val(Left(strְ��, 1))
    End If
    If intְ��B = 0 Then Exit Function '������
    
    'ȡҽ��ְ��
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "ҽ��", adVarChar, 50
        grsDuty.Fields.Append "ְ��", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.filter = "ҽ��='" & strҽ�� & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select ����,Nvl(Ƹ�μ���ְ��,0) as ְ�� From ��Ա�� Where ����='" & strҽ�� & "'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork")
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!ҽ�� = rsTmp!����
            grsDuty!ְ�� = rsTmp!ְ��
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        intְ��A = grsDuty!ְ��
    End If
        
    '���ְ��Ҫ��
    If intְ��A = 0 Then
        'ҽ��δ����ְ������
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """δ����ְ��"
    ElseIf intְ��B < intְ��A Then
        '��ֵԽСְ��Խ��
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """��ְ��Ϊ""" & Split(STR_ְ��, ",")(intְ��A - 1) & """��"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Public Function GetSysParVal(Optional ByVal int������ As Integer = -9999, Optional ByVal strDefault As String) As String
'���ܣ���ȡָ��ϵͳ������ֵ
'������int������=Ϊ-9999ʱ����ʼ��������
'      strDefault=���û��ֵ��Ϊ�յ�ȱʡֵ
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If Not grsSysPars Is Nothing Then
        If grsSysPars.State = 1 Then blnDo = False
    End If
    
    GetSysParVal = zlDatabase.GetPara(int������, glngSys)
    
'    If blnDo Then
'        strSQL = "Select ������,������,����ֵ From ϵͳ������"
'        Set grsSysPars = New ADODB.Recordset
'        Call zldatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
'    End If
'
'    If int������ <> -9999 Then
'        grsSysPars.Filter = "������=" & int������
'        If Not grsSysPars.EOF Then
'            GetSysParVal = Nvl(grsSysPars!����ֵ, strDefault)
'        Else
'            GetSysParVal = strDefault
'        End If
'    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValidPH(ByVal SerialNO As String, ByRef ErrMessage As String) As String
'���ܣ��жϱ걾�����Ƿ�Ϸ��������ظ�ʽ�������Ż������Ϣ
    Dim i As Integer, intPh As Integer, intNumber As Integer
    Dim strTmp As String, blnError As Boolean
    
    ErrMessage = "": blnError = False
    For i = 1 To Len(SerialNO)
        If InStr("0123456789", Mid(SerialNO, i, 1)) > 0 Then
            strTmp = strTmp & Mid(SerialNO, i, 1)
        ElseIf Mid(SerialNO, i, 1) = "-" Then
            If intPh = 0 Then
                '������
                If Val(strTmp) > 9999 Then
                    blnError = True
                Else
                    intPh = Val(strTmp)
                    strTmp = ""
                End If
            Else
                blnError = True
            End If
        Else
            blnError = True
        End If
        
        If blnError Then Exit For
    Next
    
    If Not blnError Then
        If intPh = 0 Then
            blnError = True
        Else
            If Val(strTmp) = 0 Or Val(strTmp) > gintNumberPH Then
                blnError = True
            Else
                intNumber = Val(strTmp)
            End If
        End If
    End If
    If blnError Then
        ErrMessage = "�걾���κŸ�ʽΪ��XXX-XXXX��" & vbCrLf & _
            "���ŷ�Χ1��9999�����ڱ�ŷ�Χ1��" & gintNumberPH
        ValidPH = ""
    Else
        ValidPH = Format(intPh, "0000") & "-" & Format(intNumber, "0000")
    End If
End Function

Public Function TransSampleNO(ByVal varSampleNO As Variant) As String
    On Error Resume Next
    
    If InStr(varSampleNO, "-") = 0 Then
        TransSampleNO = varSampleNO
    Else
        TransSampleNO = (Split(varSampleNO, "-")(0) - 1) * 10000 + Split(varSampleNO, "-")(1)
    End If
End Function

Public Function TransSampleNO_PH(ByVal varSampleNO As Variant, ByVal lngDeviceID As Long) As String
    On Error Resume Next
    Dim lngTmp As Long
    
    If lngDeviceID <> -1 Or Not gblnManualPH Or InStr(varSampleNO, "-") > 0 Then
        TransSampleNO_PH = CStr(varSampleNO)
    Else
        lngTmp = Val(varSampleNO)
        TransSampleNO_PH = Format(((lngTmp \ 10000) + 1), "0000") & "-" & Format((lngTmp Mod 10000), "0000")
    End If
End Function

Public Function GetSampleNOStr(StartNO As String, EndNO As String, Optional ByRef strErr As String) As String
    '����   ���ؿ�ʼ�ͽ����м�ı걾�ִ�
    Dim strNO As String
    Dim intRow As Integer
    Dim strTemp As String

    On Error GoTo errH

    If StartNO = "" And EndNO = "" Then Exit Function

    If StartNO = "" Then
        StartNO = EndNO
    End If

    If EndNO = "" Then
       EndNO = StartNO
    End If

    GetSampleNOStr = StartNO
    strNO = StartNO
    Do Until strNO = EndNO
        intRow = intRow + 1
        If intRow > 1000 Then
            strErr = "����걾�γ�����1000�����������룬�����걾���Ƿ���ȷ��" ' & vbCrLf & _
                    '"��ȷ�ı걾��������ĸ����ĸӦ��һ�µ��磺S1 �� S50  ����д�� S1 ��  D50"
            GetSampleNOStr = ""
            Exit Function
        End If
        
        strNO = IncStr(strNO)
        strTemp = strTemp & "," & strNO
        If Len(strTemp) > 3900 Then
            GetSampleNOStr = GetSampleNOStr & strTemp & ";"
            strTemp = ""
        End If
    Loop
    
    GetSampleNOStr = GetSampleNOStr & strTemp
    
    Exit Function
errH:
    strErr = "������(GetSampleNOStr),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Public Function IncStr(ByVal strVal As String, Optional intUpDown As Integer, Optional ByRef strErr As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
'������strVal=Ҫ��1���ַ���
'      intUpDown = 0 ��1 =1 ��1
    Dim strValuse As String
    Dim intAdd As Integer
    Dim intUp As Integer
    Dim strValue As String
    Dim strValueOne As String
    Dim strHead As String
    Dim i  As Integer
    
    On Error GoTo errH
    
    strVal = UCase(strVal)

    For i = Len(strVal) To 1 Step -1
        strValueOne = Mid(strVal, i, 1)
        If Asc(strValueOne) >= Asc("0") And Asc(strValueOne) <= Asc("9") Then
        Else
            '��������
            strHead = Mid$(strVal, 1, i)
            strVal = Mid$(strVal, i + 1)
            Exit For
        End If
    Next
    
    strVal = UCase(strVal)
    
    If intUpDown = 0 Then
        '��1
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = 1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp < 10 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "0" & strValue
                    intUp = 1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp <= Asc("Z") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "A" & strValue
                    intUp = 1
                End If
            End If
        Next
        If intUp = 1 Then
            If IsNumeric(strValueOne) Then
                strValue = "1" & strValue
            Else
                strValue = "A" & strValue
            End If
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    Else
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = -1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp >= 0 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
'                    If intAdd = 0 Then
                        strValue = "9" & strValue
'                    End If
                    intUp = -1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp >= Asc("A") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    If intAdd = 0 Then
                        strValue = "Z" & strValue
                    End If
                    intUp = -1
                End If
            End If
        Next
        If intUp = 1 Then
            strValue = -1
        End If
        If Mid(strValue, 1, 1) = "0" Or Mid(strValue, 1, 1) = "A" Then
            strValue = Mid(strValue, 2)
            If strValue = "" Then strValue = 1
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    End If
    Exit Function
errH:
    strErr = "������(IncStr),������Ϣ:" & Err.Number & " " & Err.Description
End Function

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim RptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand
    For Each rptRow In rptList.Rows
        If rptRow.Childs.Count > 0 Then rptRow.Expanded = True
    Next
    If rptList.Rows.Count < 1 Then zlReportToVSFlexGrid = False: Exit Function
        
    With vfgList
        .Clear
        .Rows = 1: .FixedRows = 1: .RowHeight(.Rows - 1) = 280
        .Cols = 0
        .MergeCells = flexMergeFree
        
        '�����и���
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = rptCol.Caption
                .ColData(.Cols - 1) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(.Cols - 1) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(.Cols - 1) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, .Cols - 1, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.Width < 20 * IIf(rptList.GroupsOrder.Count = 0, 1, rptList.GroupsOrder.Count) Then
                    .ColWidth(.Cols - 1) = 0
                Else
                    .ColWidth(.Cols - 1) = rptCol.Width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '�����и���
        Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: .RowHeight(.Rows - 1) = 280
            If rptRow.GroupRow Then
                intTiers = 0
                Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1
                    Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                .MergeRow(.Rows - 1) = True
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).Value
                Next
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0
                        Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1
                            Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptRow.Record(.ColData(lngCol)).Value
                    Else
                        .TextMatrix(.Rows - 1, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Public Function FillGrid_UQ(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '   blnMerge���Ƿ�ϲ���ͬID���С�����ϲ�����������ͬ����ֵ�ԡ�;���ָ�
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngloop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngloop) = ""
        Next
        lngRow = 0
    Else
        'Ԥ����һ����
        lngRow = objMsf.Rows - 2
    End If
    
    Do While Not rsData.EOF
        lngCurrRow = FindGridLine(objMsf, CStr(zlCommFun.Nvl(rsData("ID"))))
        If lngCurrRow = -1 Then
            lngRow = lngRow + 1
            lngCurrRow = lngRow
        
            If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
            
            On Error Resume Next
            objMsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            For lngloop = 0 To objMsf.Cols - 1
                
                If Trim(objMsf.TextMatrix(0, lngloop)) <> "" Then
                    If objMsf.TextMatrix(0, lngloop) = "#" Then
                        objMsf.TextMatrix(lngCurrRow, lngloop) = lngCurrRow
                    Else
                    
                        On Error Resume Next
                        
                        strMask = ""
                        strMask = MaskArray(lngloop)
                                                
                        On Error GoTo ErrHand
                        
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop)))
                        End If
                        objMsf.TextMatrix(lngCurrRow, lngloop) = strNewValue
                    End If
                End If
                
            Next
        End If
        
        rsData.MoveNext
    Loop
    
    FillGrid_UQ = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetLabItems(objParent As Object, Optional ByVal strType As String = "", Optional ByVal strCode As String = "", Optional ByVal lngExeDept As Long, Optional objContainer As Object = Nothing) As String
'ѡ�������Ŀ(����΢������Ŀ)���ɸ�ѡ
'strType����������
'lngExeDept��ִ�п���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim sglX As Single, sglY As Single
    Dim objMain As Object
    On Error GoTo errH
    GetLabItems = ""
    If objContainer Is Nothing Then
        Set objMain = objParent.Container
    Else
        Set objMain = objContainer
    End If
    
'    strSQL = "Select 0 As ĩ��, to_Char(Rownum) As ID, '' As �ϼ�id, ����, ���� " & _
'        "From ���Ƽ������� A " & IIf(Len(strType) = 0, "", "Where A.���� = [1] ") & _
'        "Union All " & _
'        "Select 1 As ĩ��, to_Char(B.ID) As ID, to_Char(A.ID) As �ϼ�id, B.����, B.���� " & _
'        "From (Select Rownum As ID, ���� From ���Ƽ�������) A, ������ĿĿ¼ B " & _
'        "Where B.��� = 'C' And A.���� = B.�������� " & IIf(Len(strType) = 0, "", "And B.�������� = [1]")
    
    If Len(strCode) = 0 Then
        strSQL = "Select 0 As ѡ��, B.ID, B.����, B.���� " & _
            "From ������ĿĿ¼ B Where B.��� = 'C' " & IIf(Len(strType) = 0, "", "And B.�������� = [1]")
    Else
        strSQL = "Select Distinct 0 As ѡ��, B.ID, B.����, B.���� " & _
            "From ������ĿĿ¼ B,������Ŀ���� C,���鱨����Ŀ D,������Ŀ E " & _
            "Where B.ID=C.������ĿID And B.ID=D.������ĿID " & _
            "And D.������ĿID=E.������ĿID And D.ϸ��ID Is Null And B.��� = 'C' And E.��Ŀ���<>2 " & _
            IIf(Len(strType) = 0, "", "And B.�������� = [1] ") & _
            "And (B.���� Like [2] Or C.���� Like [2] Or (Nvl(B.�����Ŀ,0)=0 And Upper(E.��д) Like [2]))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������Ŀ", strType, UCase(strCode) & "%")
    If rsTmp.EOF Then Exit Function
    Call CalcPosition(sglX, sglY, objParent)
    If rsTmp.RecordCount = 1 Then
        Call frmSelectMuli.ShowSelect(objMain, rsTmp, "����,1200,0,1;����,3000,0,1", sglX, sglY, 5000, 3000, strTitle:="������Ŀ")
        frmSelectMuli.ReturnSelect
        If rsTmp.RecordCount > 0 Then
            GetLabItems = rsTmp("ID")
        End If
        Exit Function
    End If
    
    If frmSelectMuli.ShowSelect(objMain, rsTmp, "����,1200,0,1;����,3000,0,1", sglX, sglY, 5000, 3000, strTitle:="������Ŀ") Then
        Do While Not rsTmp.EOF
            GetLabItems = GetLabItems & "," & rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        If Len(GetLabItems) > 0 Then GetLabItems = Mid(GetLabItems, 2)
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub RenumVsf(objVsf As Object, intNumCol As Integer)
    Dim lngRow As Long
    Dim intLoop As Integer
    Dim dblTmp As Double
    Const intColCount As Integer = 27
    Dim GetColCount As Integer
    
    On Error Resume Next
    
    If objVsf.Cols <= intColCount Then
        GetColCount = 0
    Else
        dblTmp = objVsf.Cols / intColCount
        If InStr(dblTmp, ".") > 0 Then
            GetColCount = Mid(dblTmp, 1, InStr(dblTmp, ".") - 1)
        Else
            GetColCount = dblTmp
        End If
    End If

    For intLoop = 0 To GetColCount
        For lngRow = 1 To objVsf.Rows - 1
            objVsf.TextMatrix(lngRow, intNumCol) = lngRow
            objVsf.Cell(flexcpData, lngRow, intLoop * intColCount, lngRow, intLoop * intColCount) = ""
        Next
    Next
End Sub
Public Function VerifyAuditingRule(lngSampleID As Long, Optional strErrMessage As String, Optional ByVal iLoadProg As Integer = 1) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                       ���ʱ������˹���
    '����                       lngSampleID �걾ID; strErrMessage ����1ʱ�Ĵ�����ʾ��iLoadProg :���ó���1-��˵��� 2-������˵���
    '����                       0 ���� 1 �н��������ʾֵ
    '
    '�����־ 3-����2-����1-������4-�쳣��5-������6-����
    '
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim int����id As Integer '
    Dim strTmp As String
    On Error GoTo errH
    
    '��������ʾֵ�Ľ��
    strSQL = " select �����־ from ����걾��¼ a , ������ͨ��� b " & _
             " Where a.ID = b.����걾id and a.id = [1] and (b.�����־ = 5 Or b.�����־ = 6)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSampleID)
    If rsTmp.EOF = False Then
        VerifyAuditingRule = 1: strErrMessage = "  ���������ʾֵ��"
    End If
    '-- �����޸ģ�������ȫ��Ϊ�գ�����ʾ��
    strSQL = "Select Count(B.ID) - Sum(Decode(Trim(b.������), Null, 1, 0)) As ���" & vbNewLine & _
             "From ����걾��¼ a , ������ͨ��� B Where a.id = b.����걾ID and  a.id = [1] and a.΢����걾 is null "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSampleID)
    Do Until rsTmp.EOF
        If Nvl(rsTmp("���")) <> "" Then
            If Val("" & rsTmp!���) <= 0 Then
               VerifyAuditingRule = 1: strErrMessage = strErrMessage & "  ���ȫ��Ϊ�գ�"
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    int����id = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
    
    'If VerifyAuditingRule <> 1 And strErrMessage = "" Then
        strSQL = "Select Zl_������˹���_Check(" & lngSampleID & "," & int����id & "," & iLoadProg & ") as ��˽�� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
        If rsTmp.RecordCount <= 0 Then
            VerifyAuditingRule = 1
            strErrMessage = strErrMessage & "  ������̵��ô���! "
            Exit Function
        End If

        If Mid(rsTmp.Fields(0).Value, 1, 2) = "1|" Then
            strTmp = "1|"
        Else
            strTmp = ""
        End If
        strErrMessage = strErrMessage & "" & Mid(rsTmp.Fields(0).Value, 15)
        strErrMessage = strTmp & strErrMessage
    'End If
    strSQL = "Zl_����걾��¼_���δͨ��(" & lngSampleID & ",'" & strErrMessage & "')"
    zlDatabase.ExecuteProcedure strSQL, "��˹���"
    If strErrMessage <> "" Then
        VerifyAuditingRule = 1
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSql = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSql, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

ErrHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function
'################################################################################################################
'## ���ܣ�  �滻Ҫ��
'## ������  BasicName     :Ҫ������
'## ���أ�  Ҫ������
'################################################################################################################
Public Function ReplaceBasic(BasicName As String, lngPatientID As Long, lngPatientPage As Integer, intPatientType As Integer, lngAdvice As Long) As String
    Dim rsTmp As New ADODB.Recordset
    gstrSql = " Select Zl_Replace_Element_Value('" & BasicName & "'," & lngPatientID & "," & IIf(lngPatientPage = 0, "Null", lngPatientPage) & _
               "," & intPatientType & "," & lngAdvice & ") From dual "
    zlDatabase.OpenRecordset rsTmp, gstrSql, gstrSysName
    ReplaceBasic = Nvl(rsTmp(0))
    
End Function
'################################################################################################################
'## ���ܣ�  ��������ָ�����ļ���ָ�����¼BLOB�ֶε�SQL���
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##         arySql()    :�ڸ����ݵĻ�������չ���ӱ����SQL��䣻��ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  �ɹ�����True��ʧ�ܷ���False
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo ErrHand
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    zlBlobSql = True
    Exit Function

ErrHand:
    Close lngFileNum
    zlBlobSql = False
End Function

Public Function GetHexStr(strChr As String) As String
    '����                   �õ����ֵ�16���Ƶ��ִ�
    '����:strChr            �����ִ�
    '����                   Rtf��ʽ���ִ�
    Dim lngloop As Long
    Dim strTmp As String
    Dim strHeight As String
    Dim strLow As String
    
    For lngloop = 1 To Len(strChr)
         If Asc(Mid(strChr, lngloop, 1)) < 0 Then
            strTmp = Hex(Asc(Mid(strChr, lngloop, 1)))
            If Len(strTmp) = 3 Then strTmp = "0" & strTmp
            strLow = Mid(strTmp, 1, 2)
            If Len(strLow) = 0 Then
                strLow = "0" & strLow
            End If
            strHeight = Mid(strTmp, 3)
            If Len(strHeight) = 0 Then
                strHeight = "0" & strHeight
            End If
            GetHexStr = GetHexStr & "\'" & strLow & "\'" & strHeight
         Else
            strTmp = Hex(Asc(Mid(strChr, lngloop, 1)))
            If Len(strTmp) = 0 Then
                strTmp = "0" & strTmp
            End If
            GetHexStr = GetHexStr & "\'" & strTmp
         End If
    Next
    If Mid(GetHexStr, 1, 1) = "\" Then
        GetHexStr = Mid(GetHexStr, 2)
    End If
    If Trim(GetHexStr) = "" Then GetHexStr = " "
End Function
Public Function GetSouceElement(RtfTxt As RichTextBox, lngElement As Long) As String
    '����                   ͨ��Ҫ�����Ƶõ�Ҫ���ִ�(���滻)
    '����: lngElement       Ҫ��Number
    '����                   Ҫ�滻��Ҫ���ִ�
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    
    strStart = "ES(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    strEnd = "EE(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    
    lngStart = InStr(RtfTxt.Text, strStart)
    lngEnd = InStr(lngStart, RtfTxt.Text, strEnd)
    
    GetSouceElement = Mid(RtfTxt.Text, lngStart, lngEnd - lngStart + 16)
    
End Function

Public Function GetReplaceElement(RtfTxt As RichTextBox, lngElement As Long, strElementReplace As String) As String
    '����                   ����Ҫ�滻��Ҫ��
    '���� lngElement        Ҫ��Number
    '     strreplace        �滻Ҫ���ִ�
    '����                   �滻��������
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strReplace As String
    Dim strNewChr As String
    
    strStart = "ES(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    strEnd = "EE(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    
    lngStart = InStr(RtfTxt.Text, strStart)
    lngEnd = InStr(lngStart, RtfTxt.Text, strEnd)
    GetReplaceElement = Mid(RtfTxt.Text, lngStart, lngEnd - lngStart + 16)
    
    lngStart = InStr(GetReplaceElement, "{")
    lngEnd = InStr(GetReplaceElement, "}")
    strReplace = Mid(GetReplaceElement, lngStart, lngEnd - lngStart + 1)
    
    GetReplaceElement = Replace(GetReplaceElement, strReplace, GetHexStr(strElementReplace))
    
    GetReplaceElement = Replace(GetReplaceElement, "\highlight2", "\highlight0")
    GetReplaceElement = Replace(GetReplaceElement, "\ulwave", "\ulnone")
        
End Function

Public Sub InstrtVerifyResult(RtfTxt As RichTextBox, lngSyllabus As Long, strSyllabusReplace As String)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                       �Ѽ��������ӵ��ļ���
    '����   lngSyllabus         ��ٱ��
    '       strSyllabusReplace
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    
    strEnd = "OE(" & Format(lngSyllabus, Replace(Space(8), " ", "0")) & ",0,0)"
    lngStart = InStr(RtfTxt.Text, strEnd)
    If lngStart = 0 Then
        strEnd = "OE(" & Format(lngSyllabus, Replace(Space(8), " ", "0")) & ",1,0)"
        lngStart = InStr(RtfTxt.Text, strEnd)
        If lngStart = 0 Then Exit Sub
    End If
    lngStart = InStr(lngStart, RtfTxt.Text, "\par") + 4
    
    RtfTxt.Text = Mid(RtfTxt.Text, 1, lngStart) & strSyllabusReplace & _
                        Mid(RtfTxt.Text, lngStart)
End Sub
Public Sub AuditingReport(RtfTxt As RichTextBox, lngSampleID As Long, intPatientType As Integer, lngPatientID As Long, intBaby As Integer, lngApplyDept As Long, _
                          lngAdviceID As Long, intRepotrCount As Integer, lngPatientPage As Integer)
    '����           ���ɼ��鱨����Ŀ
    '����           intPatientType              ������Դ
    '               lngPatientID                ����ID
    '               intBaby
    '               lngApplyDept
    '               lngAdviceID
    '               intRepotrCount
    '               lngPatientPage
    Dim rsTmp As New ADODB.Recordset
    Dim rsVerify As New ADODB.Recordset         '����ָ��
    Dim strZipFile As String                    '��ѹ�ļ���ʱ·��
    Dim strFilePath As String                   '��ʱRTF�ļ�·��
    Dim lngNewCaseHistory As Double               '�µĵ��Ӳ�����¼ID
    Dim astrSQL() As String                     '����SQL�ִ�
    Dim lngSQLCount As Integer                  'SQL�ִ����鳤��
    Dim intLoop As Integer                      '��ʱѭ������
    Dim lngResult                               '��ʾ������ID
    Dim lngNextID As Double                       '�õ���һ��ID
    Dim strSampleID As String                   '�걾ID
    Dim strLine As String                       'һ������
    Dim strRtfTxt As String                     '���ļ���������
    Dim strSouce As String                      'ԴҪ���ִ�
    Dim strReplace As String                    '�滻��Ҫ���ִ�
    Dim lngUPID As Double                         '�ϼ�ID
    Dim lngFileID As Double                       '�ļ�ID
    Dim blBeginTrans As Boolean                 '�Ƿ�ʼ����
    Dim strRtf() As String                      '����Rtf�ִ�����
    
    On Error GoTo errH
    
    If intPatientType = 1 Then
        gstrSql = "Select Count(Id) As ��ҳID From ���˹Һż�¼ Where ��¼״̬ =1 and ��¼���� =1 and  ����ID  = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngPatientID)
        If rsTmp.EOF = False Then
            lngPatientPage = Val(Nvl(rsTmp("��ҳID")))
        End If
    Else
        If intPatientType <> 2 Then
            lngPatientPage = 0
        End If
    End If
    
    gstrSql = "Select �����ļ�ID From ����ҽ����¼ a , ��������Ӧ�� b , �����ļ��б� c" & vbNewLine & _
              " Where a.������Ŀid = b.������Ŀid And b.�����ļ�id = c.Id" & vbNewLine & _
              "      And a.���Id = [1] And b.Ӧ�ó��� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngAdviceID, IIf(intPatientType = 3, 1, intPatientType))
    If rsTmp.EOF = True Then
        MsgBox "û���ҵ���Ӧ�Ĳ����ļ�ID!", vbInformation, gstrSysName
        Exit Sub
    End If
    lngFileID = rsTmp("�����ļ�ID")
    
    '������Ӳ�������
    gstrSql = "Select Id,�ļ�ID,nvl(��ID,0) as ��ID,�������,��������,������,��������,��������,�����д�,�����ı�,�Ƿ���,Ԥ�����ID" & vbNewLine & _
              "       �������,ʹ��ʱ��,����Ҫ��ID,�滻��,Ҫ������,Ҫ������,Ҫ�س���,Ҫ��С��,Ҫ�ص�λ,Ҫ�ر�ʾ,������̬,Ҫ��ֵ��" & vbNewLine & _
              " From �����ļ��ṹ Where �ļ�id = [1]" & vbNewLine & _
              "  Start With ��id  Is Null " & _
              "  Connect By Prior Id = ��id "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngFileID)
    
    '��һ��ѭ������Ҫ��
    rsTmp.filter = "�������� = 1 and �����ı� = '������'"
    If rsTmp.EOF = False Then
        lngResult = rsTmp("ID")
        rsTmp.filter = "��ID <> " & lngResult
    Else
        rsTmp.filter = ""
    End If
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "û���ҵ���Ӧ�ĵ��Ӳ�������", vbInformation, gstrSysName
        Exit Sub
    End If
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    intLoop = 0
    Do While Not rsTmp.EOF
    
        If rsTmp("ID") <> lngResult Then
            intLoop = intLoop + 1
            lngSQLCount = lngSQLCount + 1
            ReDim Preserve astrSQL(1 To lngSQLCount)
            
            lngNextID = GetNextNextId("���Ӳ�������")
            If rsTmp("��ID") = 0 Then
                lngUPID = lngNextID
            End If
            
            astrSQL(lngSQLCount) = "Zl_���Ӳ�������_Update(" & lngNextID & "," & lngFileID & ",1,0," & IIf(rsTmp("��ID") = 0, "Null", lngUPID) & "," & intLoop & "," & _
                                    rsTmp("��������") & "," & Nvl(rsTmp("������"), "Null") & "," & Nvl(rsTmp("��������"), "Null") & ",'" & _
                                    Nvl(rsTmp("��������"), "") & "'," & Nvl(rsTmp("�����д�"), "Null") & ","
                                    
            If Nvl(rsTmp("��������"), 0) = 4 And Nvl(rsTmp("�滻��"), 0) = 1 Then
                astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & _
                                       ReplaceBasic(rsTmp("Ҫ������"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID) & _
                                       "'," & Nvl(rsTmp("�Ƿ���"), "Null") & ")"
            Else
                astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & rsTmp("�����ı�") & "'," & Nvl(rsTmp("�Ƿ���"), "Null") & ")"
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '�����������µ�����
    If lngResult > 0 Then
        rsTmp.filter = "id = " & lngResult
        rsTmp.MoveFirst
        intLoop = intLoop + 1
        lngSQLCount = lngSQLCount + 1
        ReDim Preserve astrSQL(1 To lngSQLCount)
        lngNextID = GetNextNextId("���Ӳ�������")
        If rsTmp("��ID") = 0 Then
            lngUPID = lngNextID
        End If
        'д��������
        astrSQL(lngSQLCount) = "Zl_���Ӳ�������_Update(" & lngNextID & "," & lngFileID & ",1,0," & IIf(rsTmp("��ID") = 0, "Null", lngUPID) & "," & intLoop & "," & _
                                Nvl(rsTmp("��������"), "Null") & "," & Nvl(rsTmp("������"), "Null") & "," & Nvl(rsTmp("��������"), "Null") & ",'" & _
                                Nvl(rsTmp("��������"), "") & "'," & Nvl(rsTmp("�����д�"), "Null") & ",'" & Nvl(rsTmp("�����ı�")) & "'," & _
                                Nvl(rsTmp("�Ƿ���"), "Null") & ")"
                                
        'д�����ָ��
        gstrSql = "Select 0 As �������, '' As  ������Ŀ����,'   ' || rpad('������Ŀ',32) as ������Ŀ," & vbNewLine & _
                    "       rpad('���ν��', 10) As ���ν��," & vbNewLine & _
                    "       lpad('��־', 8) As ��־," & vbNewLine & _
                    "       lpad('��λ',10) As ��λ, lpad('�ο�',15) As �ο�  From dual" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select /*+ RULE */ �������, ������Ŀ����,'   ' || rpad(������Ŀ,32) as ������Ŀ," & vbNewLine & _
                    "       rpad(Decode(���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**',Null,' ', ���ν��),10,' ') As ���ν��," & vbNewLine & _
                    "       lpad(decode(��־,null,' ',��־),8,' ') as ��־," & vbNewLine & _
                    "       lpad(decode(��λ,null,' ',��λ),10,' ' ) as ��λ, lpad(decode(�ο�,Null,' ',�ο�),15,' ') as �ο���" & vbNewLine & _
                    "From (Select " & vbNewLine & _
                    "        A.������Ŀid, A.�������, A.������Ŀ As ������Ŀ����, Decode(A.�������, Null, 0, 1) As �̶���Ŀ, C.ID," & vbNewLine & _
                    "        C.������ || Decode(D.��д, Null, '', '(' || D.��д || ')') As ������Ŀ, B.ԭʼ���, '' As �ϴν��, '' As Cv," & vbNewLine & _
                    "        B.������ As ���ν��, D.���㹫ʽ, D.�������," & vbNewLine & _
                    "        Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                    "        Nvl(E.����id, -1) As ����id, Nvl(E.�걾���, 0) As �걾���, E.����ʱ��, E.�걾���," & vbNewLine & _
                    "        Decode(E.����id, Null," & vbNewLine & _
                    "                To_Char(Trunc(E.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(E.�걾���, 10000), '0000'), E.�걾���) As �걾����ʾ," & vbNewLine & _
                    "        E.���鱸ע, F.����, F.�Ա�, F.����, F.�����, F.סԺ��, F.��ǰ����, 0 As ��ҳid, D.�����Χ," & vbNewLine & _
                    "        Nvl(G.С��λ��, 2) As С��, D.��������, D.��������, D.��λ," & vbNewLine & _
                    "        Trim(Replace(Replace(' ' || Zlgetreference(C.ID, E.�걾����, Decode(E.�Ա�, '��', 1, 'Ů', 2, 0), F.��������," & vbNewLine & _
                    "                                                    E.����id, E.����,E.�������id), ' .', '0.'), '��.', '��0.')) As �ο� " & vbNewLine
                    
        gstrSql = gstrSql & "��From (Select A.������Ŀid, Min(Decode(E.������Ŀid, A.������Ŀid, F.����, 99999)) As ������Ŀ," & vbNewLine & _
                            "              Max(Decode(E.������Ŀid, A.������Ŀid, E.�������, F.����)) As �������" & vbNewLine & _
                            "       From ������ͨ��� A, ���鱨����Ŀ E, ������ĿĿ¼ F," & vbNewLine & _
                            "            (Select Distinct C.������Ŀid" & vbNewLine & _
                            "              From ������Ŀ�ֲ� B, ����ҽ����¼ C" & vbNewLine & _
                            "              Where B.�걾id = [1] And B.ҽ��id = C.���id) D" & vbNewLine & _
                            "       Where E.������Ŀid = D.������Ŀid And A.������Ŀid = E.������Ŀid(+) And E.������Ŀid = F.ID(+) And" & vbNewLine & _
                            "             A.����걾id = [1]" & vbNewLine & _
                            "       Group By A.������Ŀid" & vbNewLine & _
                            "       Order By Min(Decode(E.������Ŀid, A.������Ŀid, F.����, 99999))," & vbNewLine & _
                            "                Max(Decode(E.������Ŀid, A.������Ŀid, E.�������, F.����))) A, ������ͨ��� B, ����������Ŀ C," & vbNewLine & _
                            "     ������Ŀ D, ����걾��¼ E, ������Ϣ F, ����������Ŀ G��" & vbNewLine & _
                            "Where B.������Ŀid = A.������Ŀid(+) And B.����걾id = [1] And B.������Ŀid = C.ID And C.ID = D.������Ŀid And" & vbNewLine & _
                            "      B.����걾id = E.ID And E.����id = F.����id(+) And B.������Ŀid = G.��Ŀid(+) And B.��¼���� = [2] And" & vbNewLine & _
                            "      (G.����id = E.����id + 0 Or G.����id Is Null Or E.����id Is Null)��" & vbNewLine & _
                            "Order By ������Ŀ����, �������) A"

                            
                                    
        Set rsVerify = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngSampleID, intRepotrCount)
        
        Do While Not rsVerify.EOF
            intLoop = intLoop + 1
            lngSQLCount = lngSQLCount + 1
            ReDim Preserve astrSQL(1 To lngSQLCount)
            lngNextID = GetNextNextId("���Ӳ�������")
            
            strLine = Nvl(rsVerify("������Ŀ")) & Nvl(rsVerify("���ν��")) & Nvl(rsVerify("��λ")) & Nvl(rsVerify("��־")) & Nvl(rsVerify("�ο�"))
            'д��������
            astrSQL(lngSQLCount) = "Zl_���Ӳ�������_Update(" & lngNextID & "," & lngFileID & ",1,0," & lngUPID & "," & intLoop & ",2" & _
                                    ",Null" & "," & "Null,0,Null,'" & strLine & "',1)"
            rsVerify.MoveNext
        Loop
        
        'д������µ�����
        rsTmp.filter = " ��id = " & lngResult
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                intLoop = intLoop + 1
                lngSQLCount = lngSQLCount + 1
                ReDim Preserve astrSQL(1 To lngSQLCount)
                
                astrSQL(lngSQLCount) = "Zl_���Ӳ�������_Update(" & lngNextID & "," & lngFileID & ",1,0," & lngUPID & "," & intLoop & "," & _
                                Nvl(rsTmp("��������"), "Null") & "," & Nvl(rsTmp("������"), "Null") & "," & Nvl(rsTmp("��������"), "Null") & ",'" & _
                                Nvl(rsTmp("��������"), "") & "'," & Nvl(rsTmp("�����д�"), "Null") & ","
                                
                If Nvl(rsTmp("��������"), 0) = 4 And Nvl(rsTmp("�滻��"), 0) = 1 Then
                    astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & _
                                           ReplaceBasic(rsTmp("Ҫ������"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID) & _
                                           "'," & Nvl(rsTmp("�Ƿ���"), "Null") & ")"
                Else
                    astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & rsTmp("�����ı�") & "'," & Nvl(rsTmp("�Ƿ���"), "Null") & ")"
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    '�õ������ļ�
    strZipFile = zlBlobRead(1, lngFileID)
    If gobjFSO.FileExists(strZipFile) Then
        strFilePath = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strFilePath) = False Then
            MsgBox "û���ҵ������ļ�!", vbInformation, gstrSysName
            Exit Sub
        End If
        Kill strZipFile
    End If
    
    '����ĵ�
    RtfTxt.Text = ""
    '����Rtf�ļ�
    Open strFilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, strLine
        RtfTxt.Text = RtfTxt.Text & IIf(RtfTxt.Text <> "", vbCrLf, "") & strLine
    Loop
    Close #1
    
    '���Ҳ��滻Ҫ��
    rsTmp.filter = "�������� = 4 and �滻�� = 1 "
    Do While Not rsTmp.EOF
        strSouce = GetSouceElement(RtfTxt, rsTmp("������"))
        strReplace = GetReplaceElement(RtfTxt, rsTmp("������"), ReplaceBasic(rsTmp("Ҫ������"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID))
        RtfTxt.Text = Replace(RtfTxt.Text, strSouce, strReplace)
        rsTmp.MoveNext
    Loop
    
    'д�������
    If lngResult > 0 Then
        strReplace = ""
        If rsVerify.RecordCount > 0 Then rsVerify.MoveFirst
        Do While Not rsVerify.EOF
            strLine = Nvl(rsVerify("������Ŀ")) & Nvl(rsVerify("���ν��")) & Nvl(rsVerify("��λ")) & Nvl(rsVerify("��־")) & Nvl(rsVerify("�ο�"))
            strReplace = strReplace & "\" & GetHexStr(strLine) & "\par "
            rsVerify.MoveNext
        Loop
        
        rsTmp.filter = "�������� = 1 and �����ı� = '������'"
        rsTmp.MoveFirst
'        strRePlace = Mid(strRePlace, 5)
        InstrtVerifyResult RtfTxt, rsTmp("������"), strReplace
    End If
    
    '����RTF�ļ�
    strRtf = Split(RtfTxt.Text, vbCrLf)
    If UBound(strRtf) < 0 Then Exit Sub
    Open strFilePath For Output As #1
    For intLoop = 0 To UBound(strRtf)
        Print #1, strRtf(intLoop)     ' ���ı�����д���ļ���
    Next
    Close #1
    
'    strRtf = Split(Me.RtfTxt.Text, vbCrLf)
'    If UBound(strRtf) < 0 Then Exit Sub
'    Open "c:\10.rtf" For Output As #1
'    For intLoop = 0 To UBound(strRtf)
'        Print #1, strRtf(intLoop)     ' ���ı�����д���ļ���
'    Next
'    Close #1
    
    lngNewCaseHistory = GetNextNextId("���Ӳ�����¼")
    strZipFile = zlFileZip(strFilePath)
    If gobjFSO.FileExists(strZipFile) Then
        zlBlobSql 5, lngNewCaseHistory, strZipFile, astrSQL
    End If
    Kill strFilePath
    Kill strZipFile
        
    '--���Ӳ�����¼
    lngSQLCount = UBound(astrSQL) + 1
    ReDim Preserve astrSQL(1 To lngSQLCount)
    astrSQL(lngSQLCount) = "Zl_���Ӳ�����¼_Update(" & lngNewCaseHistory & "," & intPatientType & "," & lngPatientID & "," & lngPatientPage & "," & _
                 intBaby & "," & lngApplyDept & "," & lngFileID & "," & lngAdviceID & ")"
                 
    blBeginTrans = True
    gcnOracle.BeginTrans
    For intLoop = 1 To UBound(astrSQL)
        zlDatabase.ExecuteProcedure Replace(astrSQL(intLoop), "Call", ""), gstrSysName
'        Debug.Print aStrSQL(intLoop)
    Next
    gcnOracle.CommitTrans
    Exit Sub
errH:
    If blBeginTrans = True Then gcnOracle.RollbackTrans: blBeginTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetNextNextId(strTable As String) As Double
    '------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
    '������
    '   strTable��������
    '���أ�
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "������ü�¼" Or strtab = "סԺ���ü�¼" Then strtab = "���˷��ü�¼"
    
    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Call SQLTest(App.ProductName, "mdlCommon", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    GetNextNextId = rsTmp.Fields(0).Value

End Function

Public Function CheckExesState(lngKey As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����:      ���סԺ���˳�Ժ���Ƿ��л��۵���Ҫ�������
    '����       �걾ID
    '����       �л��۵�δ��� = Fasle û���� = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    CheckExesState = True
    
    '81��ϵͳ����Чʱ�����
'    If zlDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
    '��ǰ�����Ƿ��ѳ�Ժ��Ԥ��Ժ
    gstrSql = "select d.no" & vbNewLine & _
            "from (select distinct d.ҽ��id" & vbNewLine & _
            "       from ����걾��¼ a, ������Ϣ b, ������ҳ c, ������Ŀ�ֲ� d" & vbNewLine & _
            "       where a.����id = b.����id and a.����id = c.����id and a.��ҳid = c.��ҳid and" & vbNewLine & _
            "             a.id = [1] and a.������Դ = 2 and (b.��Ժʱ�� is not null or c.״̬ = 3) and" & vbNewLine & _
            "             a.id = d.�걾id) a, ����ҽ����¼ b, ����ҽ������ c, סԺ���ü�¼ d" & vbNewLine & _
            "where a.ҽ��id in (b.���id, b.id) and b.id = c.ҽ��id and c.��¼���� = d.��¼���� and" & vbNewLine & _
            "      c.no = d.no and d.��¼���� = 2 and d.��¼״̬ = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "���鼼ʦ����վ-����״̬���", lngKey)
    
    CheckExesState = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
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
Public Function Between(X, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function
Public Sub SendSample(WinsockC As Winsock, ByVal strIP As String, ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, _
                    Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0)
    With WinsockC
        .SendData "SendSample," & strIP & "," & lngDeviceID & "," & strSampleDate & "," & strSampleNO & "," & _
                    Replace(strAdviceIDs, ",", ";") & "," & blnUndo & "," & iType
    End With
End Sub

Public Sub GetResultFromFile(WinsockC As Winsock, ByVal strIP As String, ByVal strFile As String, ByVal lngDeviceID As Long, _
            ByVal strSampleNO As String, ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
    With WinsockC
        .SendData "ResultFromFile," & strIP & "," & strFile & "," & lngDeviceID & "," & strSampleNO & "," & dtStart & "," & _
                  dtEnd
    End With
End Sub

Public Sub Open_LIS_Report(ByVal frmParent As Object, ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng����ID As Long, ByVal lng�걾ID As Long, ByVal blnCurrMoved As Boolean, ByVal blnPrint As Boolean)
    '���ô�ͼ�ε�LIS����
    '����ͼ�ι��Զ��屨�����
'    mfrmLabMainImage.zlRefresh mlngKey, True
    Dim strChart(0 To 8) As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    On Error GoTo ErrHandle
    strSQL = "select id from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng�걾ID)
    intLoop = 0
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Debug.Print strChart(intLoop)
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                        "����ID=" & lng����ID, "�걾ID=" & lng�걾ID, _
                        "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                        "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                        "ͼ��9=" & strChart(8), IIf(blnPrint, 2, 1))
    End If

    'ɾ��ͼ���ļ�
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If Dir(strChart(intLoop)) <> "" Then Kill strChart(intLoop)
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadImageData(ByVal strPath As String, ByVal lngID As Long) As Boolean
        '�����ݿ��ȡͼ�����ݣ����ƺ󱣴浽ָ����·���¡�
        '��Σ�
        '   strPath ·��
        '   lngID   ����ͼ������ID
        '--����еĻ�, ɾ��ԭ������ʱͼ���ļ�
        Static objImg As Object
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '�Ƿ�ͼƬ��ʽ
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim objPic As New frmChartPic
        Dim killFile As String
    
        Dim blnFtp As Boolean       'FTP�Ƿ����
        Static strFtpPara As String       '����FTP����
        Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
        Dim strDownOk As String, strFtpPath   As String, strLocalFile As String
        Dim objStream As TextStream
    
        On Error GoTo ErrHandle
    
100     If Dir(strPath & "\" & lngID & ".cht") <> "" Then
102         LoadImageData = True
            Exit Function
        End If
    
        'FTP���Ӽ�飬��Ч����԰�FTP��ʽȡͼƬ
104     blnFtp = False
106     If strFtpPara = "" Then
108         strFtpPara = zlDatabase.GetPara("FTP����", glngSys, 1208, "")
        End If
110     If UBound(Split(strFtpPara, ";")) >= 3 Then
112        strFtpUser = Split(strFtpPara, ";")(0)
114        strFtpPass = Split(strFtpPara, ";")(1)
116        strFtpIP = Split(strFtpPara, ";")(2)
118        strFtpDir = Split(strFtpPara, ";")(3)
120        If TestFTP(strFtpUser, strFtpPass, strFtpIP, strFtpDir) = "" Then
122             blnFtp = True
           End If
        End If
    
124     mlngImageID = lngID
    
126     lngCount = 0
128     strFile = ""
   
130     strSQL = "select �걾id,ͼ������,ͼ��λ�� from ����ͼ���� where id = [1] "
132     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngID)
    
134     If rsTmp.EOF = True Then
            Exit Function
        End If
    
136     If objImg Is Nothing Then Set objImg = CreateObject("zlLisDev.clsDrawGraph")
    
138     Do Until rsTmp.EOF
140         strImageType = Trim("" & rsTmp("ͼ������"))
142         strFtpPath = Trim("" & rsTmp!ͼ��λ��)
144         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- ͼ��������ݿ��У���ԭ���ķ�ʽ����
146             gstrSql = "select Zl_FUN_Get����ͼ��([1],[2],[3]) from dual "
148             Set rsImage = zlDatabase.OpenSQLRecord(gstrSql, "LoadImgData", CLng(rsTmp("�걾id")), CStr(Nvl(rsTmp("ͼ������"))), CInt("0"))
150             strTmp = Nvl(rsImage(0))
152             strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
            
154             If strImageData <> "" Then
156                 intLoop = 0
                
158                 If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
160                     blnPic = True
162                     If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
164                         strFile = App.path & "\zlLisPic" & lngID & ".bmp"
166                     ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
168                         strFile = App.path & "\zlLisPic" & lngID & ".jpg"
170                     ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
172                         strFile = App.path & "\zlLisPic" & lngID & ".gif"
174                     ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
176                         If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP") = False Then
178                             gobjFSO.CreateFolder App.path & "\ZLLIS_ZIP"
                            End If
180                         If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP\" & lngID) = False Then
182                             gobjFSO.CreateFolder App.path & "\ZLLIS_ZIP\" & lngID
                            End If
184                         strFile = App.path & "\ZLLIS_ZIP\" & lngID & "\ZLISPIC.ZIP"
                        End If
                    
                    
186                     intLayOut = Val(Mid(strImageData, 1, 3))
188                     strImageData = Mid(strImageData, 5)
190                     lngFileNum = FreeFile
192                     lngCount = 0
    
194                     If Dir(strFile) <> "" Then Kill strFile
196                     Open strFile For Binary As lngFileNum
198                     ReDim aryChunk(Len(strImageData) / 2 - 1) As Byte
200                     For lngBound = LBound(aryChunk) To UBound(aryChunk)
202                         aryChunk(lngBound) = CByte("&H" & Mid(strImageData, lngBound * 2 + 1, 2))
                        Next
                    
204                     Put lngFileNum, , aryChunk()
                    
                    End If
                    '-------����ΪͼƬ�ļ�
206                 Do While strTmp <> ""
208                     intLoop = intLoop + 1
210                     gstrSql = "select Zl_FUN_Get����ͼ��([1],[2],[3]) from dual "
212                     Set rsImage = zlDatabase.OpenSQLRecord(gstrSql, "LoadImgData", CLng(rsTmp("�걾id")), CStr(Nvl(rsTmp("ͼ������"))), intLoop)
                    
214                     strTmp = Nvl(rsImage(0))
    
216                     If blnPic Then
                            '
218                         If strTmp <> "" Then
220                             ReDim aryChunk(Len(strTmp) / 2 - 1) As Byte
222                             For lngBound = LBound(aryChunk) To UBound(aryChunk)
224                                 aryChunk(lngBound) = CByte("&H" & Mid(strTmp, lngBound * 2 + 1, 2))
                                Next
                            
226                             Put lngFileNum, , aryChunk()
                            End If
                        Else
                            'ͼ������
228                         strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                        End If
                    Loop
                
230                 If blnPic Then
232                     strImageData = intLayOut & ";" & strFile
234                     Close lngFileNum
                    End If
                End If
            Else
                'ͼ�����FTP�У���FTP��ȡ����
                'ͼ��λ�õ����ݸ�ʽΪ��ͼ���ʽ;FTP�ļ�·��
            
236             intLayOut = Val(Split(strFtpPath, ";")(0))
238             strFtpPath = Trim(Split(strFtpPath, ";")(1))
240             strImageData = ""
242             If intLayOut >= 100 And intLayOut <= 227 Then
                    ' ͼƬ�ļ���ֱ�����ص�����
244                 strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
246                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
248                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
250                 If strDownOk = "" Then
252                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' ͼ�����ݣ���Ҫ�����ص��ı��ļ��ж�ȡ����
254                 strLocalFile = strPath & "\" & lngID & "_" & strImageType & ".txt"
256                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
258                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
260                 If strDownOk = "" Then
262                     Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
264                     Do Until objStream.AtEndOfLine
266                         strImageData = strImageData & objStream.ReadLine
                        Loop
268                     objStream.Close
270                     Set objStream = Nothing
272                     strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
274                     strImageData = intLayOut & ";" & strImageData
                    End If
276                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                End If
            End If
        
278         If Len(strImageData) <> 0 Then
280             If Not objImg Is Nothing Then
282                 LoadImageData = objImg.DrawImg(strImageType, strImageData, strPath & "\" & lngID & ".cht")
                End If
            End If
        
284         strTmp = "": strImageData = ""
286         rsTmp.MoveNext
        Loop
        Exit Function
ErrHandle:
        WriteLog "mdlLisWork", "LoadImagedata", CStr(Erl()) & "�У�" & Err.Description
288     If ErrCenter() = 1 Then
290         Resume
        End If
End Function

Public Function ReadVerifyData(lngID As Long, intRule As Integer) As String
    ''''''''''''''''''''''''''''''''''''''''''''''
    '����       ���ɵ���ǩ�����õ��ִ�
    '����       lngID=�걾ID
    '           intRule=�����ִ�����
    '����       ���ɺõ��ִ�
    '''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim strData As String
    Dim strBase As String
    
    On Error GoTo errH
    If intRule = 1 Then
        '�õ�������Ϣ
        gstrSql = "Select ID, ҽ��id, �걾���, ������, ����ʱ��, �걾����, ������, ����ʱ��, ������, ����ʱ��, �����, ���ʱ��, ��������," & vbNewLine & _
                "       ����id, ��������, ������, ��ע, ����ʱ��, �걾��̬, ִ�п���id," & vbNewLine & _
                "       ΢����걾, NO, �걾���, ���鱸ע, ������, �������id, ������Դ, ����id, Ӥ��, ����, �Ա�, ��������, ���䵥λ," & vbNewLine & _
                "       ����, �Һŵ�, �����, סԺ��, ��������, ��ҳid, ������Ŀ, ��������," & vbNewLine & _
                "       ������, ����ʱ��, ��ʶ��, ����, ���˿���,  ����" & vbNewLine & _
                "From ����걾��¼ A" & vbNewLine & _
                "Where A.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����ǩ��", lngID)
        If rsTmp.EOF = True Then Exit Function
        For intLoop = 0 To rsTmp.Fields.Count - 1
            strBase = strBase & "," & rsTmp(intLoop)
        Next
        strBase = Mid(strBase, 2)
        '�õ����
        gstrSql = "Select ����걾id, ������Ŀid, ������, �����־, ����ο�, �޸���, �޸�ʱ��, ��¼����, ԭʼ���, ԭʼ��¼ʱ��, ��¼��," & vbNewLine & _
                "       �Ƿ����, �޸�ԭ��, ϸ��id, ����id, ��������, ������Ŀid, �������, Od, Cutoff, Sco, ø���id, ���ý��" & vbNewLine & _
                "From ������ͨ��� " & vbNewLine & _
                "where ����걾ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����ǩ��", lngID)
        Do Until rsTmp.EOF
            strData = strData & "|"
            For intLoop = 0 To rsTmp.Fields.Count - 1
                strData = strData & "," & rsTmp(intLoop)
            Next
            rsTmp.MoveNext
        Loop
        strData = Mid(strData, 3)
        ReadVerifyData = strBase & ";" & strData
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Signature(lngID As Long, Optional strAuditingMan As String, Optional strType As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           ��LIS���浥ǩ��
    '����           lngID=����걾ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSource As String                                 'ȡ�õ��Ӳ���ǩ���ִ�
    Dim lng֤��ID As Long                                   '֤��ID
    Dim strSign As String                                   'ǩ�������ɵ��ִ�
    Dim strTimeStamp As String                              'ʱ���
    Dim strTimeStampCode As String                          'ʱ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim intSaveInfoSign As Integer                          '���յǼǱ���ʱǩ�� 1=����
    Dim intSaveReprotSign As Integer                        '���浥����ʱǩ�� 0=����
    Dim strSQL As String
    
    
    '��鵱ǰ�����Ƿ�ʹ��ǩ��
    strSQL = "select ִ�п���ID from ����걾��¼ where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����ǩ��", lngID)
    
    strSQL = "select Zl_Fun_Getsignpar(6," & rsTmp("ִ�п���ID") & ") as tag from dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����ǩ��")
    
    If rsTmp("tag") = 0 Then
        'û�����õĿ���ֱ�Ӵ���Ϊǩ���ɹ�
        Signature = True
        Exit Function
    End If
    
    
    intSaveInfoSign = zlDatabase.GetPara("���յǼǱ���ʱǩ��", 100, 1208, 1)
    intSaveReprotSign = zlDatabase.GetPara("���浥����ʱǩ��", 100, 1208, 1)
    
    If strType = "����" And intSaveInfoSign = 0 Then
        Signature = True
        Exit Function
    End If
    
    If strType = "����" And intSaveReprotSign = 0 Then
        Signature = True
        Exit Function
    End If
    
    On Error GoTo errH
    '����ǩ��
    If Not gobjESign Is Nothing Then
        If Not gobjESign.CheckCertificate(IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser)) Then Exit Function
        If gobjESign.CertificateStoped(UserInfo.����) = False Then
            strSource = ReadVerifyData(lngID, 1)
            If strSource = "" Then
                MsgBox "���ܶ�ȡҪǩ���ļ��鱨�浥��", vbInformation, gstrSysName
                Exit Function
            End If
            strSign = gobjESign.Signature(strSource, IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser), lng֤��ID, strTimeStamp, , strTimeStampCode)
            If strSign = "" Then Exit Function
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            strTimeStampCode = IIf(strTimeStampCode = "", "NULL", "'" & strTimeStampCode & "'")
            gstrSql = "Select A.���� From ��Ա�� A, �ϻ���Ա�� B Where A.Id = B.��Աid and b.�û��� = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "����ǩ��", IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser))
            If rsTmp.EOF = False Then
                gstrSql = "zl_����ǩ����¼_Insert(" & lngID & ",1,'" & Replace(strSign, "'", "''") & _
                             "'," & lng֤��ID & "," & strTimeStampCode & "," & strTimeStamp & ",'" & rsTmp("����") & "')"
                zlDatabase.ExecuteProcedure gstrSql, "����ǩ��"
            End If
        End If
    End If
    Signature = True
    Exit Function
errH:
    Err.Raise Err.Number, "����ǩ��"
End Function

Public Function VerifySignature(lngID As Long) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           ��֤ǩ��
    '����           lngID = ����걾ID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    VerifySignature = gobjESign.VerifySignature(ReadVerifyData(lngID, 1), lngID, 4)
        
    
End Function
Public Function GetAdviceMoney(ByVal str��ID As String, ByVal strҽ��ID As String, ByVal str���ͺ� As String, _
    str��� As String, str����� As String, Optional ByVal bln����ִ�� As Boolean, Optional ByVal strItemType As String) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      strItemType=�Ƿ�ʹ���������������ƣ���Ҫ�������ֲɼ�
'���أ�str���,str�����=���ڱ�����ʾ
'˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strSQLbak As String
    Dim intPatientType As Integer                                   '������Դ
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ ������Դ From ����ҽ����¼ Where ID In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str��ID)
    If rsTmp.EOF = True Then Exit Function
    intPatientType = rsTmp("������Դ")
    
    If bln����ִ�� Then
        strSQL = _
            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From סԺ���ü�¼ A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 = [2] And (mod(A.��¼����,10), A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
            
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
        End If
    Else
'        strSQL = _
'            " Select B.����,B.����,Sum(A.ʵ�ս��) as ��� From סԺ���ü�¼ A,�շ���Ŀ��� B,����ҽ����¼ C" & _
'            " Where A.ҽ����� + 0 In" & _
'            "      (Select ID From ����ҽ����¼" & _
'            "       Where ID In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)" & _
'            "       Union All" & _
'            "       Select ID From ����ҽ����¼" & _
'            "       Where ���id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B))" & _
'            "  And (mod(A.��¼����,10) , A.NO) In" & _
'            "      (Select ��¼����, NO From ����ҽ������" & _
'            "       Where ҽ��id In" & _
'                "      (Select ID From ����ҽ����¼" & _
'                "       Where ID In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)" & _
'                "       Union All" & _
'                "       Select ID From ����ҽ����¼" & _
'                "       Where ���id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B))" & _
'            "         And ���ͺ� + 0 In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B)" & _
'            "       Union All" & _
'            "       Select ��¼����, NO From ����ҽ������" & _
'            "       Where ҽ��id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B)" & _
'            "         And ���ͺ� + 0 In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B))" & _
'            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.���� and a.ҽ����� = c.id  "
        strSQL = "Select *" & vbNewLine & _
                "From (With T1 As (Select /*+cardinality(b,10)*/ b.Column_Value As ID" & vbNewLine & _
                "                  From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B" & vbNewLine & _
                "                  Union All" & vbNewLine & _
                "                  Select /*+cardinality(b,10)*/ a.Id" & vbNewLine & _
                "                  From ����ҽ����¼ A, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B" & vbNewLine & _
                "                  Where a.���id = b.Column_Value)" & vbNewLine & _
                "       Select b.����, b.����, Sum(a.ʵ�ս��) As ���" & vbNewLine & _
                "       From סԺ���ü�¼ A, �շ���Ŀ��� B, ����ҽ����¼ C, T1" & vbNewLine & _
                "       Where a.ҽ����� = T1.Id And (Mod(a.��¼����, 10), a.No) In" & vbNewLine & _
                "             (Select /*+cardinality(b, 10)*/ a.��¼����, a.No" & vbNewLine & _
                "              From ����ҽ������ A, Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B, T1" & vbNewLine & _
                "              Where a.ҽ��id = T1.Id And a.���ͺ� = b.Column_Value" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select /*+cardinality(a, 10) cardinality(b, 10)*/ a.��¼����, a.No" & vbNewLine & _
                "              From ����ҽ������ A, Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B," & vbNewLine & _
                "                   Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) C" & vbNewLine & _
                "              Where a.ҽ��id = b.Column_Value And a.���ͺ� = c.Column_Value) And a.���ʷ��� = 1 And a.��¼״̬ = 0 And" & vbNewLine & _
                "             a.�շ���� = b.���� And a.ҽ����� = c.Id "
        strSQL = strSQL & _
            "" & IIf(strItemType <> "", " And c.������� = [5] ", "") & _
            " Group by B.����, B.����)"
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
        End If
    End If
'    strSQLbak = strSQL
'    strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
'    strSQL = strSQL & " union all " & strSQLbak
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str��ID, strҽ��ID, str���ͺ�, glngSys, strItemType)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer, Optional ByVal bln���� As Boolean) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
'     str�շ����=��ǰҪ�������,���ڷ��౨��
'     str�������=�������,������ʾ
'     bln����=���ɻ��۷���ʱ�ı��������ƾ���ǿ�Ƽ���Ȩ��ʱ�Ĵ���
'     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
'����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
'     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
'     0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    If mintWarn = 0 Then mintWarn = intWarn
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - cur���ʽ��
    cur���ս�� = cur���ս�� + cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then mintWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then mintWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If mintWarn = 0 Then
                                BillingWarn = 2
                            ElseIf mintWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & " ����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then mintWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then mintWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If mintWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf mintWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                                If mintWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then mintWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If mintWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                    If vMsg = vbIgnore Then mintWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then mintWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then mintWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If mintWarn = 0 Then
                                BillingWarn = 2
                            ElseIf mintWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function
Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rspati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If lng��ҳID <> 0 Then
        'סԺ���˱���
        strSQL = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
        
        strSQL = "Select A.����,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rspati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, lng��ҳID)
        zlDatabase.Currentdate
    Else
        '���������ﱨ��
        strSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1]"
        strSQL = "Select A.����,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
            " From ������Ϣ A,(" & strSQL & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
            " Where A.����ID=B.����ID(+) And A.����id = D.����id(+) And A.����=D.����(+)" & _
            " And D.����=E.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
            " And A.����ID=[1]"
        Set rspati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID)
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strSQL = "Select Nvl(��������,1) as ��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, CStr(Nvl(rspati!���ò���)))
    If Not rsWarn.EOF Then
        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(lng����ID)
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rspati!����), Nvl(rspati!ʣ���, 0), cur����, cur���, Nvl(rspati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Chk���۷���(Objfrm As Object, strҽ��ID�� As String, lng�걾ID As Long, Optional strItemType As String) As Boolean
    '���� ���黮�۵�����ʱ�Ǳ��ѽ�ֹ��������
    '���� strҽ��ID��=����ҽ��ʱֱ�Ӵ���ҽ������","�ָ�
    '     lng�걾ID=���걾ID�����Ҫ��ҽ��ID
    '     �Ƿ�Ȩ��������Ŀ�����
    '     ������������ֻ��һ��������
    Dim curMoney As Currency, str��� As String, str����� As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strIDs As String
    Dim int������Դ As Integer
    Dim int��ҳID As Integer
    Dim lng����ID As Long
    Dim strҽ��ID As String
    Dim str���ͺ� As String
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    If lng�걾ID <> 0 Then
        strSQL = "Select Distinct Decode(B.ҽ��id, Null, A.ID, B.ҽ��id) As ҽ��id" & vbNewLine & _
                "From ����걾��¼ A, ������Ŀ�ֲ� B" & vbNewLine & _
                "Where A.ID = B.�걾id And A.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Chk���۷���", lng�걾ID)
        Do While rsTmp.EOF
            strIDs = strIDs & "," & rsTmp("ҽ��ID")
            rsTmp.MoveNext
        Loop
        strIDs = Mid(strIDs, 2)
    Else
        strIDs = strҽ��ID��
    End If
    
    strIDs = Replace(Replace(strIDs, ";", ","), "|", ",")
    
    strSQL = "Select /*+ rule */ a.����ID,A.��ҳid, A.������Դ, B.ҽ��id, B.���ͺ�, C.��ǰ����id" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C" & vbNewLine & _
            "Where A.ID = B.ҽ��id And A.����id = C.����id And" & vbNewLine & _
            "      A.ID In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Chk���۷���", strIDs)
    If rsTmp.EOF = True Then Exit Function
    int������Դ = Nvl(rsTmp("������Դ"))
    int��ҳID = Nvl(rsTmp("��ҳID"), 0)
    lng����ID = Nvl(rsTmp("��ǰ����id"), 0)
    lng����ID = Nvl(rsTmp("����ID"))
    Do While Not rsTmp.EOF
        strҽ��ID = strҽ��ID & "," & rsTmp("ҽ��ID")
        str���ͺ� = str���ͺ� & "," & rsTmp("���ͺ�")
        rsTmp.MoveNext
    Loop
    curMoney = GetAdviceMoney(strҽ��ID, strҽ��ID, str���ͺ�, str���, str�����, False, strItemType)
    
    If Not FinishBillingWarn(Objfrm, gstrPrivs, lng����ID, int��ҳID, lng����ID, curMoney, str���, str�����) Then Exit Function
    
    Chk���۷��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function
Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    'û��ǿ�������ٴ�,����ҽ��������
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
    If bln���� Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
            " Where A.����ID=B.����ID And A.��ԱID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----������ FTP ��غ���
Private Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP��
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.path, 1) <> "\", App.path & "\", App.path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP����" & strPath) > 0 Then
            TestFTP = "��FTP�ϲ��ܴ���Ŀ¼��"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "�ϴ��ļ�ʧ��"
            Else
                FtpNet.FuncFtpDisConnect '�ȶϿ�����ɾ������Ȼɾ����
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP�������ӣ�"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP����" & strPath) > 0 Then
                    TestFTP = "��FTP�ϲ���ɾ��Ŀ¼"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "��������FTP��"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Private Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '��FTP�����������ļ���
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpFile :FTP�ϵ��ļ���
        'strFile    :�����ļ�ȫ·����
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
        Dim objFtp As New clsFtp, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "��ָ��Ҫ���ص��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "��ָ�����ص��ļ����浽�δ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "Ҫ���ص��ļ��Ѵ��ڣ�"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "�������ӷ�������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFile = "���ܽ���ָ����Ŀ¼��������Ȩ�޲������������޴�Ŀ¼��"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFtpDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "����ʧ�ܣ�������Ȩ�޲������������޴��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "�У�" & Err.Description
End Function

Private Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '�������ļ����ϴ��ļ���FTP��������
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpPath :FTP�ϵ�Ŀ¼����Ŀ¼���Զ�������
        'strFile    :�����ļ�ȫ·����
        'strNewFileName: ����FTP�Ϻ���ļ�����Ϊ���򰴱����ļ�������
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
    
        Dim objFtp As New clsFtp, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "�ļ�" & strLocaFile & "������!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '���Ŀ¼�Ƿ����
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "����Ŀ¼ʧ�ܣ�������Ȩ�޲��㣡"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "�ϴ��ļ�ʧ�ܣ�������Ȩ�޲��㣡"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "�������ӷ�������"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "�У�" & Err.Description
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

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function


Public Sub WriterBarCodeToLIS(rsBarcode As ADODB.Recordset, intMode As Integer, Optional ByVal intContinue As Integer)
    '����   ������д��LIS���뵥
    Dim strErr As String
    If rsBarcode.RecordCount > 0 Then
        If Not mobjLisInsideComm Is Nothing Then
            rsBarcode.MoveFirst
            Do Until rsBarcode.EOF
                If intMode = 3 Then
                    If mobjLisInsideComm.SampleBarcodeWrite(rsBarcode("ҽ��ID��"), rsBarcode("��������"), UserInfo.����, strErr, intContinue) = False Then
                        MsgBox "д�����뵽LIS���뵥����!" & vbCrLf & strErr
                    End If
                Else
                    If mobjLisInsideComm.SampleBarcodeWrite(rsBarcode("ҽ��ID��"), "", "", strErr, intContinue) = False Then
                        MsgBox "д�����뵽LIS���뵥����!" & vbCrLf & strErr
                    End If
                End If
                rsBarcode.MoveNext
            Loop
        End If
    End If
End Sub
Public Sub WriterSampleSendDateToLIS(strAdvice As String, intType As Integer, ByVal strUser As String)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����   ���ͼ�ʱ��д��LIS���뵥��
    '����   strAdvice = ҽ���������ŷָ�
    '       intType = 0 �ͼ�  1 = ȡ���ͼ�
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleSendInfo(strAdvice, intType, strUser, strErr) = False Then
            MsgBox "д���ͼ�ʱ�䵽LIS���뵥����!" & vbCrLf & strErr
        End If
    End If
End Sub

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strWidth As String, strText As String, i As Integer
        
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To objThis.Cols - 1
        strWidth = strWidth & "," & objThis.Body.ColWidth(i)
        If UCase(TypeName(objThis)) = UCase("BillEdit") Then
            If objThis.msfObj.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        Else
            If objThis.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        End If
    Next
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "���", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "����", Mid(strText, 2)
    
    If TypeName(objThis) = "VSFlexGrid" Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "����", objThis.FrozenCols
    End If
End Sub
Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim strWidth As String, strText As String
    Dim arrText As Variant, i As Integer
        
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        RestoreFlexState = True: Exit Function
    End If
    
    
        
    strWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "���", "")
    If UBound(Split(strWidth, ",")) >= objThis.Cols - 1 Then
        For i = 0 To objThis.Cols - 1
            objThis.Body.ColWidth(i) = Split(strWidth, ",")(i)
        Next
        RestoreFlexState = True
    End If
    
    
End Function

Public Function CheckDocEmpower(ByVal lng������ĿID As Long, ByVal strAppend As String) As Boolean
'���ܣ�������Ա�Ƿ����������Ŀ��ִ��Ȩ
'������strAppend=��ǰ���븽�����д�����,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    strSQL = "select A.ID from ����������Ŀ A,������������ B where a.����id=b.id and b.����='06' and A.������='����ҽ��'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpower")
    If rsTmp.RecordCount > 0 Then
        lngID = rsTmp!ID
        arrItem = Split(strAppend, "<Split1>")
        For i = 0 To UBound(arrItem)
            arrSub = Split(arrItem(i), "<Split2>")
            If Val(arrSub(2)) = lngID Then
                If Trim(arrSub(3)) <> "" Then
                    strDoc = Trim(arrSub(3))
                End If
                Exit For
            End If
        Next
    End If
    If strDoc = "" Then strDoc = UserInfo.����
    strSQL = "Select Count(*) as Ȩ�� From ��Ա����Ȩ�� A,��Ա�� B Where A.��Աid = B.ID And B.����=[1] And A.������Ŀid = [2] And A.��¼���� = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng������ĿID)
    CheckDocEmpower = Val(rsTmp!Ȩ�� & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetOrderInspectInfo(ByVal lng����ID As Long, ByVal strCondition As String) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵ�
    
    If gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = DynamicCreate("zl9EmrInterface.ClsEmrInterface", "�°没��")
    End If
    If Not gobjEmrInterface Is Nothing Then
        GetOrderInspectInfo = gobjEmrInterface.GetOrderInspectInfo(lng����ID, strCondition)
    End If
    
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function BlnIsNumber(ByVal strCode As String) As Boolean
    '���֣��������ж�
     If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        BlnIsNumber = True
     Else
        BlnIsNumber = False
     End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "") As String
    '����:����Ϸ��Լ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay))
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = Nvl(rsTemp.Fields(0).Value)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ModifyApplyToLIS(strAdvices As String, intType As Integer)
    '����   ��ǩ����Ϣд��LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.ModifyApplyItemStateYJ(strAdvices, intType) = False Then
            
        End If
    End If
End Sub

Public Function GetAdvicePrice(ByVal lngPatientID As Long, ByVal lngPageID As Long) As String
    Dim strYPJG As String   'ҩƷ�۸�ȼ�
    Dim strWCJG As String   '���ļ۸�ȼ�
    Dim strPTXM As String   '��ͨ��Ŀ�۸�ȼ�
    
    On Error GoTo ErrHand
    
    If gobjpublicExpenses Is Nothing Then
        Set gobjpublicExpenses = CreateObject("zlPublicExpense.clsPublicExpense")
        Call gobjpublicExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    If gobjpublicExpenses.zlGetPriceGrade(gstrNodeNo, lngPatientID, lngPageID, "", strYPJG, strWCJG, strPTXM) = True Then
        If strPTXM <> "" Then
            GetAdvicePrice = " = '" & strPTXM & "'"
        Else
            GetAdvicePrice = " is null "
        End If
    Else
        GetAdvicePrice = " is null "
    End If
    
    Exit Function
ErrHand:
    MsgBox "������(GetAdvicePrice),������Ϣ:" & Err.Number & " " & Err.Description, vbInformation, "��ʾ"
    Err.Clear
        
End Function
