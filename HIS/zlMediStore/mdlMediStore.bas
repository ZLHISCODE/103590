Attribute VB_Name = "mdlMediStore"
Option Explicit
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrprivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrStockSearchPrivs As String       'ר����Կ���ѯ��Ȩ��

Public glngModul As Long
Public glngSys As Long                      'ϵͳ��Ų���
Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrSQL As String                    '������Ϊ������ʱSQL���
Public gstrDbUser As String                 '��ǰ��¼ORACLE�û���
Public mblnCostPrice As Boolean             '���ⵥ���Ƿ���ʾ�ɱ���
Public Const GCST_INVALIDCHAR = "'"             '�����������Ч�ַ�

Public Const StrFormat As String = "'999999999990.99999'"
Public gstrMatchMethod As String            'ƥ�䷽ʽ:0��ʾ˫��ƥ��
Public gstrUserName As String               '�����û�����
Public gobjDrugPurchase As Object           '�ɹ�ƽ̨����
Public gbytSimpleCodeTrans As Byte          '��Ƭ�����Ƿ���������л�����

'�û���Ϣ------------------------
Public Type TYPE_USER_INFO
    �û�ID As Long
    �û����� As String
    �û����� As String
    �û����� As String
    ����ID As Long
    ���ű��� As String
    �������� As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum �༭
    '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
    '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������;
    '8-ҩ���˻�
    ���� = 1
    �޸� = 2
    ��� = 3
    ���� = 4
    �޸ķ�Ʊ = 5            '���������˵ĵ��ݽ��й�ҩ��λ����Ʊ��Ϣ�����޸�
    ���� = 6
    ������� = 7            '���ڶ�����˵ĵ��ݽ��гɱ��ۡ���ҩ��λ����Ʊ��Ϣ����ˣ�����ԭʼ���ݣ������µ��ݣ�
    ҩ���˻� = 8            '����ҩ���򹩻���λ�˻�
    �˲� = 9                '���ں˲�ɱ���
    ���� = 10               '�����³���ⷿ�Ŀ�������
End Enum

'ҩƷ����ѯ�У������α�����������ɫ����
Public Const glng���� As Long = &HC00000
Public Const glng���� As Long = &H80000008
Public Const glngͣ�� As Long = &HC0

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012
 
Public gint���뷽ʽ As Integer              '0-ƴ����1-���
Public gintҩƷ������ʾ As Integer          '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
Public gint����ҩƷ��ʾ As Integer          '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��

Public grsMaster As New ADODB.Recordset        'ҩƷѡ������ҩƷ��񻺴����ݼ�
Public grsMasterInput As New ADODB.Recordset   'ҩƷѡ������ҩƷ���¼�����ʱ�Ļ������ݼ�
Public grsSlave As New ADODB.Recordset         'ҩƷѡ���������λ������ݼ�

Public gstrPriceClass As String         '�۸�ȼ�

'ģ���
Public Enum ģ���
    �⹺��� = 1300
    ������� = 1301
    ������� = 1302
    ��۵��� = 1303
    ҩƷ�ƿ� = 1304
    ҩƷ���� = 1305
    �������� = 1306
    ҩƷ�̵� = 1307
    ҩƷ�ƻ� = 1330
    �������� = 1331
    ҩƷ���� = 1333
End Enum

'ҵ�񵥾ݺ�
Public Enum ���ݺ�
    �⹺��� = 1
    ������� = 2
    Эҩ��� = 3
    ������� = 4
    ��۵��� = 5
    ҩƷ�ƿ� = 6
    ҩƷ���� = 7
    �շѴ�����ҩ = 8
    ���ʵ�������ҩ = 9
    ���ʱ�����ҩ = 10
    �������� = 11
    �̵�� = 12
    ���۱䶯 = 13
    �̵㵥 = 14
    �����¼ = 27
End Enum


'ҩƷ��ͨģ��Ҫʹ�õ���ϵͳ����
Public Type Type_SysParms
    P9_���ý���λ�� As Integer
    P29_ָ�������۶��۵�λ As Integer
    P44_����ƥ�� As String
    P54_ʱ��ҩƷ�ԼӼ������ As Integer
    P64_������� As Integer
    P75_�⹺�����Ҫ�˲� As Integer
    P76_ʱ��ҩƷֱ��ȷ���ۼ� As Integer
    P85_ҩ���鿴���ݳɱ��� As Integer
    P96_ҩƷ��¿��ÿ�� As Integer
    P126_ʱ��ҩƷ�ۼۼӳɷ�ʽ As Integer
    P149_Ч����ʾ��ʽ As Integer
    P150_ҩƷ���������㷨 As Integer
    P173_������Ǹ������ܽ��и������ As Integer
    P174_ҩƷ�ƿ���ȷ���� As Integer
    P175_ҩƷ������ȷ���� As Integer
    P181_ҩƷ��ⰴ�ֶμӳ� As Integer
    P183_ʱ��ȡ�ϴ��ۼ� As Integer
    P221_ҩƷ���ʱ�� As Integer
    P275_���۹���ģʽ As Integer
    P294_����ȡĿ¼�в�����Ϣ As Integer
End Type
Public gtype_UserSysParms As Type_SysParms     'ϵͳ����

'ҩƷ���۸�������󾫶�
Public Type Type_Digits
    Digit_��� As Integer
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

Public Type Type_SaleDigits
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserSaleDigits As Type_SaleDigits

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
     x As Long
     y As Long
End Type

'API����
Public Const GWL_HWNDPARENT = (-8)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Function ExistsColObject(Col, index) As Boolean
    '�жϼ������Ƿ����ָ������(�ؼ���)�ĳ�Ա
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    If TypeName(Col(index)) = "Collection" Then
        '������Ӧ�ĳ�Ա�Ǽ���ʱ
        ExistsColObject = True
        Exit Function
    Else
        '������Ӧ�ĳ�Ա�ǷǼ���ʱ
        v = Col(index)
        ExistsColObject = True
        Exit Function
    End If
ErrorHandler:
    '�쳣ʱ��ʾ��������Ӧ�ĳ�Ա
    ExistsColObject = False
End Function
Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '�����չ�ӿڳ�ʼ��
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub


Public Sub zlPlugIn_SetVBMenu(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, FrmMain As Form)
    '������չ���ܵĲ˵���Ŀ������VB�Դ��˵�����ҪԼ��zlPlugIn���˵�����Ϊ"mnuPlugIn"���Ӳ˵�������Ϊ"mnuPlugItem"
    '������lngSys-ϵͳ��lngModul-ģ��ţ�objPlugIn-��չ��Ҷ���FrmMain-���ڶ���
    Dim strFunc As String, strFuncName As String '��¼��չ����
    Dim blnGroup As Boolean
    Dim i As Integer
    Dim intCount As Integer
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '��Ҳ�������չ����
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        
        Err.Clear: On Error GoTo 0
    End If
    
    If strFunc = "" Then Exit Sub
    
    FrmMain.mnuPlugIn.Visible = True
    
    strFunc = Replace(strFunc, "Auto:", "")

    For i = 0 To UBound(Split(strFunc, ","))
        strFuncName = Split(strFunc, ",")(i)
        blnGroup = InStr(strFuncName, "|") > 0
        strFuncName = Replace(strFuncName, "InTool:", "")
        strFuncName = Replace(strFuncName, "|:", "")
        
        If i <> 0 Then
            If blnGroup Then
                '�зָ�ʱ���ٲ���һ���ָ��˵�
                intCount = intCount + 1
                Load FrmMain.mnuPlugItem(intCount)
                FrmMain.mnuPlugItem(intCount).Caption = "-"
            End If
            
            intCount = intCount + 1
            Load FrmMain.mnuPlugItem(intCount)
        End If
        
        FrmMain.mnuPlugItem(intCount).Caption = strFuncName
        FrmMain.mnuPlugItem(intCount).Tag = strFuncName
        
        If i <= 9 Then
            FrmMain.mnuPlugItem(intCount).Caption = strFuncName & "(&" & IIf(i = 9, 0, i + 1) & ")"
        End If
    Next
End Sub

Public Sub zlPlugIn_Fun(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, FrmMain As Form, _
    ByVal strFunName As String, ByVal strParams As String)
    '������չ���ܲ˵�����ִ��
    '������lngSys-ϵͳ��lngModul-ģ��ţ�objPlugIn-��չ��Ҷ���FrmMain-���ڶ���
    '      strFunName-��������,strParams-���ܲ���(��ʽ���ⷿid,����,NO)
    Dim lng�ⷿID As Long
    Dim int���� As Integer
    Dim strNo As String
    
    On Error Resume Next
    
    lng�ⷿID = Val(Split(strParams, ",")(0))
    int���� = Val(Split(strParams, ",")(1))
    strNo = Split(strParams, ",")(2)
    
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        Call objPlugIn.DrugStuffWorkNoramal(lngModul, strFunName, lng�ⷿID, strNo, int����)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� ExecuteFunc ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Sub zlPlugIn_SetVBToolbar(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, _
    FrmMain As Form, tlbTool As Toolbar, strPlugInKey As String, strPlugInSeparatorKey As String)
    '������չ���ܵĹ�������Ŀ������VB�Դ��ؼ�
    '������lngSys-ϵͳ��lngModul-ģ��ţ�objPlugIn-��չ��Ҷ���cbrToolBar-CommandBar����������lngMenuPlugInMain-��Ҳ˵�
    Dim strFunc As String, strFuncName As String '��¼��չ����
    Dim blnGroup As Boolean
    Dim i As Integer
    Dim intKeyIndex As Integer  '��ťkeyֵ�Զ��������
    Dim intIndex As Integer '��ť����
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '��Ҳ�������չ����
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        
        Err.Clear: On Error GoTo 0
    End If

    If strFunc = "" Then Exit Sub

    For i = 0 To UBound(Split(strFunc, ","))
        strFuncName = Split(strFunc, ",")(i)
        
        '���ݸ�ʽ���빤������ť
        If InStr(strFuncName, "InTool:") > 0 Then
            blnGroup = InStr(strFuncName, "|") > 0
            strFuncName = Replace(strFuncName, "InTool:", "")
            strFuncName = Replace(strFuncName, "|:", "")
            
            With FrmMain.tlbTool.Buttons
                If intIndex = 0 Then
                    'PlugIn��ť����
                    intIndex = .Item(strPlugInKey).index
                End If
                
                '��ʾPlugIn��ʼ�ָ���ť
                .Item(strPlugInSeparatorKey).Visible = True
                
                If i = 0 Then
                    '��һ�����ܰ�ť�Ѵ��ڣ���ʾ����
                    .Item(strPlugInKey).Visible = True
                Else
                    If blnGroup = True Then
                        '���ӷָ���ť
                        .Add intIndex + 1, "PlugItem" & intKeyIndex + 1, strFuncName, 3
                        intIndex = intIndex + 1
                        intKeyIndex = intKeyIndex + 1
                    End If
                    
                    '���Ӹ����PlugIn���ܰ�ť
                    .Add intIndex + 1, "PlugItem" & intKeyIndex + 1, strFuncName, 0, .Item(strPlugInKey).Image
                    intIndex = intIndex + 1
                    intKeyIndex = intKeyIndex + 1
                End If
            End With
        End If
    Next
End Sub


Public Sub zlPlugIn_Unload(objPlugIn As Object)
    'ж����ҽӿ�
    Set objPlugIn = Nothing
End Sub
Public Function Get�ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get�ۼ�-ȡ����ҩƷ�ۼ�", lngҩƷID)
        
        If Not rsData.EOF Then
            Get�ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        gstrSQL = "select Decode(���ۼ�, Null, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
            " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            '��������ݣ��Ӽ۸��ȡ
            gstrSQL = "Select �ּ� As ���ۼ� From ҩƷ�۸��¼ Where �۸����� = 1 And ��¼״̬ = 1 And ҩƷid = [1] And �ⷿid = [2] And nvl(����,0) = [3] "
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷID, lng�ⷿID, lng����)
             
            If Not rsData.EOF Then
                Get�ۼ� = rsData!���ۼ�
            Else
                '�۸�������ݣ��ӹ����ȡ���һ�μ۸�
                gstrSQL = "Select �ϴ��ۼ� as ���ۼ�,ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID = [1] "
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷID)
                
                If Not rsData.EOF Then
                    If Not IsNull(rsData!���ۼ�) Then
                        '���ϴ��ۼ�ȡֵ
                        Get�ۼ� = rsData!���ۼ�
                    Else
                        '���ϴ��ۼ�ʱ�����ݳɱ��ۼ�����е����ݼ���
                        'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
                        '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                        '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                        dblָ�����ۼ� = rsData!ָ�����ۼ�
                        dbl��������� = rsData!���������
                        
                        Get�ۼ� = 0
                        dbl�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng����)
                        dbl�ӳ��� = rsData!�ӳ��� / 100
                        dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                        dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                        Get�ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
                    End If
                End If
            End If
        Else
            '���������
            Get�ۼ� = rsData!���ۼ�
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function CheckIsAccount(ByVal lng�ⷿID As Long) As Boolean
'    '�ж��Ƿ����Ѿ��������Ѿ����
'    Dim rsData As ADODB.Recordset
'    Dim lng���id As Long
'
'    gstrSQL = "Select Nvl(Max(ID), 0) as ���id From ҩƷ����¼ Where �ⷿid = [1] "
'    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsAccount", lng�ⷿID)
'
'    lng���id = rsData!���ID
'
'    '���֮ǰ���й����
'    If lng���id > 0 Then
'        gstrSQL = "Select �ڳ�����, ��ĩ����, ������, ��������, �����, �������, �ϴν��id, �ڼ�, ���� From ҩƷ����¼ Where id=[1]"
'        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsAccount", lng���id)
'
'        '����Ƿ���δ��˵Ľ������
'        If Not rsData.EOF Then
'            If Nvl(rsData!�������) = "" Then
'                MsgBox "��ʾ��������ݻ�δ��ˡ�" & vbCrLf & "Ϊȷ������׼ȷ�ԣ�������˽�棬�ٽ�������ҵ�������", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'    End If
'
'    CheckIsAccount = True
'End Function

Public Sub AutoAdjustPrice_ByID(ByVal lngDrugID As Long)
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '��ָ��ҩƷID���
    '��ҩƷѡ�����е���
    
    On Error GoTo errHandle
    
    gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngDrugID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice_ByID")

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_ByNO(ByVal int���� As Integer, ByVal strNo As String)
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '��ָ������,NO�е�ҩƷ�Ž��м��
    '����ͨҵ��ģ������ʱ����
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.ҩƷid " & _
        " From �շѼ�Ŀ A, ҩƷ�շ���¼ B, �շ���ĿĿ¼ C " & _
        " Where a.�շ�ϸĿid = b.ҩƷid And a.�շ�ϸĿid = c.Id And Nvl(c.�Ƿ���, 0) = 0 And a.�䶯ԭ�� = 0 And a.ִ������ <= Sysdate And b.������� Is Null " & _
        " And b.���� = [1] And b.No = [2]" & GetPriceClassString("A") & _
        " Union " & _
        " Select Distinct a.ҩƷid " & _
        " From ҩƷ�۸��¼ A, ҩƷ�շ���¼ B " & _
        " Where a.ҩƷid = b.ҩƷid And a.��¼״̬ = 0 And a.ִ������ <= Sysdate And b.������� Is Null And " & _
        " b.���� = [1] And b.No = [2] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice", int����, strNo)

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("��������ִ�е��ۣ����Ժ�......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !ҩƷID
            gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngAdjustID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_Batch()
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '�������ҩƷ
    '��ҩƷѡ�������ݼ���ʼ��ʱ����
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct a.�շ�ϸĿid As ҩƷid" & vbNewLine & _
        "From �շѼ�Ŀ A, �շ���ĿĿ¼ B" & vbNewLine & _
        "Where a.�շ�ϸĿid = b.Id And b.��� In ('5', '6', '7') And Nvl(b.�Ƿ���, 0) = 0 And a.�䶯ԭ�� = 0 " & _
        "And a.ִ������ <= Sysdate" & GetPriceClassString("A") & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct a.ҩƷid From ҩƷ�۸��¼ A Where a.��¼״̬ = 0 And a.ִ������ <= Sysdate"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice")

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("��������ִ�е��ۣ����Ժ�......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !ҩƷID
            gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngAdjustID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNotVerifyClosingAccount() As ADODB.Recordset
    '��ѯ��ǰ����Ա�����Ĳ����Ƿ����δ��˵Ľ���¼
    Dim rsData As ADODB.Recordset
    Dim strDept As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.Id, b.����, 'δ������' As ����" & vbNewLine & _
            "From ������Ա A, ���ű� B, ��������˵�� C, ҩƷ����¼ D, ҩƷ������ E" & vbNewLine & _
            "Where a.����id = b.Id And b.Id = c.����id And b.Id = d.�ⷿid And d.Id = e.���id And a.��Աid = [1] And" & vbNewLine & _
            "      c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���') And d.������� Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Distinct b.Id, b.����, 'δ��˽��' As ����" & vbNewLine & _
            "From ������Ա A, ���ű� B, ��������˵�� C" & vbNewLine & _
            "Where a.����id = b.Id And b.Id = c.����id And a.��Աid = [1] And c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���') And" & vbNewLine & _
            "      Exists (Select 1 From ҩƷ����¼ D Where b.Id = d.�ⷿid And d.������� Is Null)"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "����ѯ", UserInfo.�û�ID)
    
    Set CheckNotVerifyClosingAccount = rsData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'ȡҩƷ���۸��������С��λ��
Public Function GetDigitTiaoJia(ByVal int��� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���1-ҩƷ;2-����
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ҩƷ��λ:1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '����: 0-������;1-��ʾ����
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax��� As Integer
    Dim intMax�ɱ��� As Integer
    Dim intMax���ۼ� As Integer
    Dim intMax���� As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum = 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ����")
    
    intMax��� = rs.Fields(0).NumericScale
    intMax�ɱ��� = rs.Fields(1).NumericScale
    intMax���ۼ� = rs.Fields(2).NumericScale
    intMax���� = rs.Fields(3).NumericScale
    
    If int���� = 4 Then
        int��λ = 5
    End If
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ��� = [1] And ���� = [2] And ��λ = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int���, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        GetDigitTiaoJia = rsTmp!����
    End If
    
    If GetDigitTiaoJia = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        GetDigitTiaoJia = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����, intMax���)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigitTiaoJia = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����, intMax���)
End Function

Public Function IsPriceAdjustMod(ByVal lngҩƷID As Long) As Boolean
    '�ж�ҩƷ�Ƿ��������۹���
    Dim rsData As ADODB.Recordset
    
    If gtype_UserSysParms.P275_���۹���ģʽ = 0 Then Exit Function
    
    gstrSQL = "Select Nvl(�Ƿ����۹���, 0) As �Ƿ����۹��� From ҩƷ��� Where ҩƷid = [1] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsPriceAdjustMod", lngҩƷID)
    
    If rsData.EOF Then IsPriceAdjustMod = False: Exit Function
    
    IsPriceAdjustMod = (rsData!�Ƿ����۹��� = 1)
End Function

Public Function CheckPriceAdjust(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '�޿��ʱ���ɱ���ȡҩƷ���ĳɱ���
    '������lngҩƷid-ҩƷ���ID��Ϊ0��������ҩƷ��lng�ⷿid-��Ӧ�ĿⷿID��Ϊ0�������пⷿ��lng����-��Ӧ�����Σ��������-1�򲻹�������
    '���أ�True-������false-�в��������۹���Ҫ���ҩƷ
    '
    Dim rsData As ADODB.Recordset
    Dim str���� As String
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjust = True: Exit Function
    
    '������޿��
    If lngҩƷID > 0 Then
        If lng�ⷿID > 0 Then
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] and �ⷿid=[2] " & _
                " And Not (nvl(����,0) = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
            
            If lng���� > 0 Then
                gstrSQL = gstrSQL & " and Nvl(����,0)=[3] "
            End If
        Else
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] " & _
                " And Not (nvl(����,0) = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
        End If
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            '�޿��ʱ�����շѼ�Ŀȡ�ۼۣ���ҩƷ���ȡ�ɱ��ۣ����Ƚϼ۸�
            gstrSQL = "Select a.�ɱ���, b.�ּ� As �ۼ� " & _
                " From ҩƷ��� A, �շѼ�Ŀ B " & _
                " Where a.ҩƷid = b.�շ�ϸĿid And (Sysdate Between b.ִ������ And b.��ֹ����) And Nvl(a.�Ƿ����۹���, 0) = 1 " & _
                " And b.�ּ� <> a.�ɱ��� And a.ҩƷid = [1] " & GetPriceClassString("B")
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷID)
            
            If rsData.EOF Then
                'û�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = True
            Else
                '�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = False
            End If
            
            Exit Function
        End If
    End If
    
    If lngҩƷID > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and a.ҩƷid=[1] "
    End If
    
    If lng�ⷿID > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and d.�ⷿid=[2] "
    End If
    
    If lng���� >= 0 Then
        str���� = IIf(str���� = "", "", str����) & " and nvl(d.����,0)=[3] "
    End If
    
    gstrSQL = "Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ���� " & vbNewLine & _
        "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "             b.�ּ� <> nvl(d.ƽ���ɱ���,a.�ɱ���) " & str���� & GetPriceClassString("B") & vbNewLine & _
        "  And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) " & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ���� " & vbNewLine & _
        " From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
        " Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And nvl(d.���ۼ�,0) <> nvl(d.ƽ���ɱ���,a.�ɱ���) " & str���� & _
        "  And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷID, lng�ⷿID, lng����)
    
    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then CheckPriceAdjust = True: Exit Function
    
    CheckPriceAdjust = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub


Public Function Get�ּ�(ByVal lngҩƷID As Long) As Double
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "[��ȡ��ҩƷ�����۵�λ�۸�]", lngҩƷID)
    
    If Not rsTemp.EOF Then
        Get�ּ� = rsTemp!�ּ�
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
    Dim strSQL As String
    Dim rsUser As New ADODB.Recordset
    
    Set rsUser = Sys.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            UserInfo.�û�ID = !id '��ǰ�û�id
            UserInfo.�û����� = !��� '��ǰ�û�����
            UserInfo.�û����� = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            UserInfo.�û����� = IIf(IsNull(!����), "", !����)  '��ǰ�û�����
            UserInfo.����ID = !����ID '��ǰ�û�����id
            UserInfo.���ű��� = !������ '��ǰ�û�
            UserInfo.�������� = !������  '��ǰ�û�
            UserInfo.strMaterial = GetMaterial(UserInfo.����ID)
            GetUserInfo = True
        Else
            UserInfo.�û�ID = 0 '��ǰ�û�id
            UserInfo.�û����� = "" '��ǰ�û�����
            UserInfo.�û����� = "" '��ǰ�û�����
            UserInfo.�û����� = "" '��ǰ�û�����
            UserInfo.����ID = 0    '��ǰ�û�����id
            UserInfo.���ű��� = ""  '��ǰ�û�
            UserInfo.�������� = ""  '��ǰ�û�
        End If
    End With
End Function

Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub
Public Function CheckRepeatMedicine(ByVal MyBill As Object, ByVal strDrugInfo As String, ByVal intExceptRow As Integer) As Boolean
    'ҩƷ��ͨ�༭������¼���ҩƷ�Ƿ��ظ�
    'MyBill�����ؼ���ҩƷ�б�
    'strDrugInfo��ҩƷID�����μ���Ӧ�кţ���ʽ��ҩƷID,ҩƷID�к�|����,�����кţ�
    'intExceptRow���ų�ָ�����У��������һ�У�
    Dim n As Integer
    Dim lngҩƷID As Long
    Dim intҩƷID�к� As Integer
    Dim lng���� As Long
    Dim int�����к� As Integer
    
    lngҩƷID = Val(Split(Split(strDrugInfo, "|")(0), ",")(0))
    intҩƷID�к� = Val(Split(Split(strDrugInfo, "|")(0), ",")(1))
    lng���� = Val(Split(Split(strDrugInfo, "|")(1), ",")(0))
    int�����к� = Val(Split(Split(strDrugInfo, "|")(1), ",")(1))
    
    With MyBill
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If n <> intExceptRow And Val(.TextMatrix(n, intҩƷID�к�)) = lngҩƷID And Val(.TextMatrix(n, int�����к�)) = lng���� Then
                    If MsgBox("�Բ������и�ҩƷ���ҩƷ����ͬ���Σ������ظ����룡 ��Ҫ�ƶ���������" _
                        , vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                        .Row = n
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRepeatMedicine = True
End Function

Public Function GetCheck�ⷿ(ByVal lng�ⷿID As Long) As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ƿ���������", lng�ⷿID)
    If Not rsTemp.EOF Then GetCheck�ⷿ = NVL(rsTemp!�����, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CheckStopMedi(ByVal varInput As Variant)
    '���ҩƷ�Ƿ�ͣ��
    'varInput���ָ�ʽ�����뵥����Ϣ������|No��;����ҩƷID������ʽ��ҩƷID1��ҩƷID2.....��
    Dim rsTemp As ADODB.Recordset
    Dim strMsg As String
    Dim int���� As Integer
    Dim strNo As String
    Dim n As Integer
    Dim strҩ�� As String
    
    On Error GoTo errHandle
    If InStr(varInput, "|") > 0 Then
        int���� = Mid(varInput, 1, InStr(varInput, "|") - 1)
        strNo = Mid(varInput, InStr(varInput, "|") + 1)
        
        gstrSQL = "Select Distinct '[' || C.���� || ']' AS ҩƷ����,C.���� As ͨ����,B.���� As ��Ʒ�� " & _
                " From ҩƷ�շ���¼ A, �շ���Ŀ���� B, �շ���ĿĿ¼ C " & _
                " Where A.ҩƷid = C.ID And A.ҩƷid = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                " And Nvl(C.����ʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') " & _
                " And A.���� = [1] And A.NO = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ͣ��ҩƷ", int����, strNo)
    Else
        gstrSQL = "Select Distinct '[' || C.���� || ']' AS ҩƷ����,C.���� As ͨ����,B.���� As ��Ʒ�� " & _
                " From Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) A, �շ���Ŀ���� B, �շ���ĿĿ¼ C " & _
                " Where A.Column_Value = C.ID  And A.Column_Value = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                " And Nvl(C.����ʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ͣ��ҩƷ", varInput)
    End If
    
    With rsTemp
        If Not .EOF Then
            For n = 1 To .RecordCount
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = !ҩƷ���� & !ͨ����
                Else
                    strҩ�� = !ҩƷ���� & IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
                End If
                
                If n > 5 Then
                    strMsg = strMsg & vbCrLf & "��������" & .RecordCount - 5 & "��ҩƷ......"
                    Exit For
                End If
                strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & strҩ��
                .MoveNext
            Next
            
            strMsg = "ע�⣬����ҩƷ�ѱ�ͣ�ã�" & vbCrLf & strMsg
        End If
    End With
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNoStock(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long, Optional ByVal lng���� As Long = -1) As Boolean
    '����Ƿ��޿�棬�����ж�ʱ�۲�����ҩƷ�޿���̵�ʱ����
    '���ʱ�������Σ�ֻ����������
    '���أ�true-�޿��;false-�п��
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From ҩƷ��� " & _
        " Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And (Nvl(ʵ������, 0) <> 0 Or Nvl(ʵ�ʽ��, 0) <> 0 Or Nvl(ʵ�ʲ��, 0) <> 0) "
    
    If lng���� <> -1 Then
        gstrSQL = gstrSQL & " And Nvl(����,0) = [3] "
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckNoStock", lng�ⷿID, lngҩƷID, lng����)
    
    CheckNoStock = rsData.EOF
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNumStock(ByVal objVSF As Object, ByVal lng�ⷿID As Long, ByVal lntColҩƷid As Integer, ByVal intCol���� As Integer, ByVal intCol���� As Integer, ByVal intCol����ϵ�� As Integer, ByVal intMethod As Integer, Optional int���ҵ�� As Integer, Optional ByVal int���� As Integer) As String
    '���ܣ���˳����൥��ʱ��������ʵ�������Ƿ��㹻
    '������objVSF-��Ҫ���ı��;lng�ⷿid��intcol����-���������У�intCol����-���������У�intCol����ϵ��-����ϵ��������
    '������intMethod��1-������ˣ�2-������3-�˿����
    '������int���ҵ��0-��⣻1-����
    '����ֵ�����о����ҩƷ���ƣ�Ϊ��-���ͨ�����������㣻��Ϊ��-���δͨ��������������
    Dim objCol As Collection         '��ʹ�õ���������
    Dim i, j As Integer
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim rsData As ADODB.Recordset
    Dim strKey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    Dim dbl����ϵ�� As Double
    
    '����ϱ�������������������Ҫ�ǿ��ǲ����������
    Set objCol = New Collection
    With objVSF
        If .rows < 2 Then Exit Function
        For lngRow = 1 To .rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntColҩƷid) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntColҩƷid) & "," & Val(.TextMatrix(lngRow, intCol����)) & "," & Val(.TextMatrix(lngRow, intCol����ϵ��)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strKey = .TextMatrix(lngRow, lntColҩƷid) & "," & Val(.TextMatrix(lngRow, intCol����)) & "," & Val(.TextMatrix(lngRow, intCol����ϵ��))
                '����С��λ�����������������ʱ�����������ݱȽ�
                strArray = dblNum + (Val(.TextMatrix(lngRow, intCol����)))
                objCol.Add Array(strKey, strArray), strKey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '��ʽ��ҩƷid,����,����ϵ��
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int���ҵ�� = 0 Then '���
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = Abs(dblNum)
                Else
                    '������⣬������棬���Բ����
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '�˿���ˣ��˿����¼������
                dblNum = dblNum
            End If
        Else    '����
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = 0
                Else
                    '������⣬������棬���Բ����
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        'ֻ�����������ж�
        If dblNum > 0 Then
            For i = 0 To UBound(varTemp)
                lngҩƷID = varTemp(0)
                lng���� = varTemp(1)
                dbl����ϵ�� = varTemp(2)
'                int���� = Len(Split("" & dblNum & ".", ".")(1))
                
                gstrSQL = "Select a.ʵ������, '[' || b.���� || ']' || b.���� ����" & vbNewLine & _
                            "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                            "Where a.ҩƷid = b.Id And a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] And b.��� In ('5', '6', '7') And a.���� = 1"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", dblNum, lngҩƷID, lng�ⷿID, lng����)
                If rsData.RecordCount = 0 Then
                    gstrSQL = "select '[' || ���� || ']' || ���� ���� from �շ���ĿĿ¼ where id=[1]"
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷID)
                    
                    CheckNumStock = rsData!����
                    Exit Function
                Else
                    If zlStr.FormatEx(rsData!ʵ������ / dbl����ϵ��, int����, , False) >= dblNum Then
                        CheckNumStock = ""
                    Else
                        
                        CheckNumStock = rsData!����
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Public Function ���ʵ���������(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal Dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal lngС��λ�� As Long) As Boolean
'���ܣ��ڳ���ʱ�����ʵ�������Ƿ��㹻���㹻�򷵻�true����֮Ϊfalse
    Dim rsData As ADODB.Recordset
    Dim str���� As String
    
    On Error GoTo errHandle
    
    '������޿��
    If lngҩƷID <= 0 Then Exit Function
    If lng�ⷿID <= 0 Then Exit Function
    
    gstrSQL = "Select a.ʵ������, '[' || b.���� || ']' || b.���� ����" & vbNewLine & _
                            "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                            "Where a.ҩƷid = b.Id And a.ҩƷid = [1] And a.�ⷿid = [2] And Nvl(a.����, 0) = [3] And b.��� In ('5', '6', '7') And a.���� = 1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷID, lng�ⷿID, lng����)
    
    If rsData.RecordCount = 0 Then '�޿���¼
        gstrSQL = "select '[' || ���� || ']' || ���� ���� from �շ���ĿĿ¼ where id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷID)
        
        ���ʵ��������� = False
        Exit Function
    Else '�п���¼
        ���ʵ��������� = zlStr.FormatEx(rsData!ʵ������ / dbl����ϵ��, lngС��λ��, , True) >= Dbl���� 'ʵ���������ڳ�������
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckUsableNum( _
    ByVal lng�ⷿID As Long, _
    ByVal lngҩƷID As Long, _
    ByVal lng���� As Long, _
    ByVal dbl��д���� As Double, _
    ByVal dbl����ϵ�� As Double, _
    ByVal strNo As String, _
    ByVal int���� As Integer, _
    ByVal int����� As Integer, _
    ByVal int�������� As Integer, _
    Optional int��� As Integer, _
    Optional dblSum As Double) As Boolean
    '������д����ʱ���������������Ƿ��㹻����������/�޸ģ����������
    '����ֵ true-ͨ����飬false-û��ͨ�����
    '��Σ�dbl��д�����ǽ��浥λ����
    '      strNo="", ��-� �ǿ�-�޸ģ��޸�ʱ��Ҫ�ų���ǰ��������
    '      dblSum �����ҩƷ����д�����������ڳ���/�������ʱ
    '1.���δ���0�ǰ����μ�飬����=0���Ǳ�ʾ�������飻�޸�״̬ʱҪ����ԭ����������������Ҫ���ǿ��ܱ�����δ�������ηֽ��ҵ��ռ�õ�����
    '2.�������Ҫ�����ľͲ��õ��ú�������������
    '3.����/�ƿⵥ�ݳ���ʱ���⴦��:
    '�������ȡԭ�������Σ�ע��Ҫ��ԭ������ⷿ(����ʱΪ���ⷿ)������ʱ��֧�ֶԳ���������޸ģ����Բ��������е��ݵ������Ҫ�ӽ��洫��������
    '4.���ѻ��ֹʱ���ݷ���������������������ͬ
    Dim dblNum As Double
    Dim rsData As ADODB.Recordset
    Dim dblCheck As Boolean
    Dim bln�������� As Boolean
    Dim bln�������� As Boolean
    Dim strSqlStock As String, strSqlStockBatch As String  '����������������ͷ�������
    Dim strSqlSum As String, strSqlSumBatch As String      '���ϲ�δ��˵��������������ͷ�������
    Dim lng�������� As Long
    Dim blnNewNo As Boolean '�Ƿ���������
    Dim dbl����д���� As Double
    
    On Error GoTo errHandle
    
    If int����� = 0 Then CheckUsableNum = True: Exit Function

    If int���� = 6 And int��� > 0 Then
        blnNewNo = True
        
        'ȡԭ�����Ǳʵ�����
        gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where  " & _
            " �ⷿid=[1] And ���� = [2] And NO = [3] And ��� = [4] And ҩƷid = [5] And ���ϵ�� = 1"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", lng�ⷿID, int����, strNo, int��� + 1, lngҩƷID)
        
        If rsData.RecordCount = 0 Then Exit Function
        
        lng�������� = rsData!����
        
        If lng�������� = 0 Then
            '��������Ϊ��������������������
            dbl����д���� = dblSum
        Else
            '��������Ϊ����������������ε���д����
            dbl����д���� = dbl��д����
        End If
    Else
        blnNewNo = (strNo = "")
        lng�������� = lng����
        dbl����д���� = dbl��д����
    End If
        
    strSqlStock = "Select Sum(Nvl(��������, 0)) As �������� From ҩƷ��� Where ����=1 And �ⷿid = [1] And ҩƷid = [2]"
    strSqlStockBatch = "Select Sum(Nvl(��������, 0)) As �������� From ҩƷ��� Where ����=1 And �ⷿid = [1] And ҩƷid = [2] And nvl(����,0) = [3] "
    strSqlSum = "Select Sum(��������) As ��������" & vbNewLine & _
                " From (Select Nvl(��������, 0) As ��������" & vbNewLine & _
                "       From ҩƷ���" & vbNewLine & _
                "       Where ����=1 And �ⷿid = [1] And ҩƷid = [2] " & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Abs(a.ʵ������ * Nvl(a.����, 1)) As ��������" & vbNewLine & _
                "       From ҩƷ�շ���¼ A" & vbNewLine & _
                "       Where a.������� Is Null And a.�ⷿid = [1] And a.ҩƷid + 0 = [2]  And a.No = [4] And a.���� = [5])"
    strSqlSumBatch = "Select Sum(��������) As ��������" & vbNewLine & _
                    " From (Select Nvl(��������, 0) As ��������" & vbNewLine & _
                    "       From ҩƷ���" & vbNewLine & _
                    "       Where ����=1 And �ⷿid = [1] And ҩƷid = [2]  And nvl(����,0) = [3] " & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Abs(a.ʵ������ * Nvl(a.����, 1)) As ��������" & vbNewLine & _
                    "       From ҩƷ�շ���¼ A" & vbNewLine & _
                    "       Where a.������� Is Null And a.�ⷿid = [1] And a.ҩƷid + 0 = [2]  And a.No = [4] And a.���� = [5]  And nvl(����,0) = [3] )"
    
    If lng���� = 0 Then
        '1.�����������
        If blnNewNo = True Then
            '1.1����ǵ�������״̬��ֱ�ӿ�����ܿ��������Ƿ��㹻
            gstrSQL = strSqlStock
        Else
            '1.2����ǵ����޸�״̬��Ҫ�ϲ�ԭ��������
            gstrSQL = strSqlSum
        End If
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��������", lng�ⷿID, lngҩƷID, lng��������, strNo, int����)
        
        If NVL(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl����ϵ��, int��������, True, False)
        End If
        
        If dblNum < dbl����д���� Then
            dblCheck = True
            bln�������� = True
        End If
    Else
        '2.���������
        If blnNewNo = True Then
            '2.1����ǵ�������״̬��ֱ�ӿ�����ܿ��������Ƿ��㹻
            gstrSQL = strSqlStockBatch
        Else
            '2.2����ǵ����޸�״̬��Ҫ�ϲ�ԭ��������
            gstrSQL = strSqlSumBatch
        End If
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��������", lng�ⷿID, lngҩƷID, lng��������, strNo, int����)

        If NVL(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl����ϵ��, int��������, True, False)
        End If
        
        If dblNum < dbl����д���� Then
            '2.2.1��������
            dblCheck = True
            bln�������� = True
        End If
    End If
        
    '��治��ʱ���ѻ��ֹ
    If dblCheck = True Then
        gstrSQL = "select ����,���� from �շ���ĿĿ¼ where id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷID)
                    
        Select Case int�����
        Case 1  '��ʾ
            If int���� = 2 Then '�������
                If bln�������� = True Then
                    If MsgBox("���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln�������� = True Then
                    If MsgBox("���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Else
                If bln�������� = True Then
                    If MsgBox("��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln�������� = True Then
                    If MsgBox("��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Case 2  '��ֹ
            If int���� = 2 Then '�������
                If bln�������� = True Then
                    MsgBox "���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã����ܳ��⣡", vbInformation, gstrSysName
                ElseIf bln�������� = True Then
                    MsgBox "���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "�����ܳ��⣡", vbInformation, gstrSysName
                End If
            Else
                If bln�������� = True Then
                    MsgBox "��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã����ܳ��⣡", vbInformation, gstrSysName
                ElseIf bln�������� = True Then
                    MsgBox "��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "�����ܳ��⣡", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End Select
    End If
    CheckUsableNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get��������(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long) As Integer
    '����ָ���ⷿ��ָ��ҩƷ�ķ�������
    '���أ�0-��������1-����
    Dim rsCheck As New ADODB.Recordset
    Dim int���� As Integer
    Dim blnҩ�� As Boolean
    Dim strSQL As String
        
    On Error GoTo errHandle
    
    '�ж��Ƿ���ҩ�����Ƽ���
    strSQL = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lng�ⷿID)

    blnҩ�� = (Not rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    strSQL = " Select Nvl(ҩ�����,0) As ҩ�����,nvl(ҩ������,0) As ҩ������ " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lngҩƷID)
              
    If blnҩ�� Then
        int���� = rsCheck!ҩ������
    Else
        int���� = rsCheck!ҩ�����
    End If
    
    Get�������� = int����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckStrickUsable(ByVal int���� As Integer, ByVal lng�ⷿID As Long, _
        ByVal lngҩƷID As Long, ByVal strҩƷ���� As String, _
        ByVal lng���� As Long, ByVal dbl�������� As Double, ByVal int����� As Integer, _
        Optional ByVal strNo As String = "", Optional ByVal int��� As Integer = 0) As Boolean
    '��������ʱ��飺ԭ�������ⷿ�Ƿ���������㹻�������������ڻ�С��ʵ����������ʵ�ʳ����������ܴ��ڿ�������
    '�����ƿⵥ�ݡ�����ⵥ����Ҫȡԭ��������Ǳʵ����Σ��ٸ���������ȡ����������
    '����������⡢Э����ⵥ�ݣ�������ȫ�����������Ը��ݵ��ݺţ������ȡ���������������Ϳ����������Ƚ�
    '�������ݿ�ֱ�Ӹ�������ȡ����������
    'int����飺��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    'ֻ�г���ʱ�ǳ������ͣ�ԭ������������ͣ���Ҫ���˼�飺�⹺��⡢������⣨ԭ��������Ǳʣ���Э����⣨ԭ��������Ǳʣ���������⡢�ƿ⣨ԭ��������Ǳʣ�
    
    Dim rsTemp As ADODB.Recordset
    Dim lng������� As Long
    Dim dbl�������� As Double
    
    On Error GoTo errHandle
    '��������Ϊ0ʱ���Բ���ҪУ�����������ų�����Ϊ����������ɿ���������С��0�������޷������������
    If dbl�������� = 0 Then
        CheckStrickUsable = True
        Exit Function
    End If
    
    If int���� = 2 Or int���� = 3 Then  '������⡢Э����ⵥ��
        If strNo = "" Or int��� = 0 Then Exit Function
        gstrSQL = "Select 1 From ҩƷ�շ���¼ A, ҩƷ��� B " & _
            " Where A.���� = [1] And A.NO = [2] And A.��� = [3] And A.��¼״̬ = 1 And A.���ϵ�� = 1 And B.���� = 1 And A.�ⷿid = B.�ⷿid And A.ҩƷid = B.ҩƷid And " & _
            " Nvl(A.����, 0) = Nvl(B.����, 0) And A.ʵ������ > B.ʵ������ And Rownum = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", int����, strNo, int���)
        
        '���������̽�����ʾ���ֹ
        If rsTemp.RecordCount > 0 Then
            Select Case int�����
            Case 1  '��ʾ
                If MsgBox(strҩƷ���� & "�Ŀ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '��ֹ
                MsgBox strҩƷ���� & "�Ŀ�治�㣡", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    Else
        If int���� = 6 Or int���� = 4 Then   '�ƿⵥ��������ⵥ
            If strNo = "" Or int��� = 0 Then Exit Function
            
            gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��� = [3] And ҩƷid = [4] And ���ϵ�� = 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", int����, strNo, int���, lngҩƷID)
            
            If rsTemp.RecordCount = 0 Then Exit Function
            
            lng������� = rsTemp!����
        Else
            '�������ݸ��ݴ����������ȡ����������
            lng������� = lng����
        End If
        
        gstrSQL = "Select Nvl(ʵ������, 0) ʵ������ From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿID, lngҩƷID, lng�������)
        
        If rsTemp.RecordCount > 0 Then
            dbl�������� = rsTemp!ʵ������
        End If
        
        '���������̽�����ʾ���ֹ
        If dbl�������� < Abs(dbl��������) Then
            Select Case int�����
            Case 1  '��ʾ
                If MsgBox(strҩƷ���� & "�Ŀ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '��ֹ
                MsgBox strҩƷ���� & "�Ŀ�治�㣡", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    End If
    
    CheckStrickUsable = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetControlItem(ByVal int���� As Integer, ByVal int���� As Integer) As String
    '���ݻ��ڿ��ƣ������޸ĵ���Ŀ����ʱֻ���⹺���
    'int���ڣ�1-�˲�;2-���;3-������ˣ�ҩƷ�⹺��
    '������Ŀ���ɹ���,����,�����,������,�ۼ�,���,��Ʊ��,��Ʊ����,��Ʊ���
    Dim rsTmp As ADODB.Recordset
    Dim strControlItem As String
    Const cst����_�⹺ As Integer = 1
    Const cst����_�˲� As Integer = 1
    Const cst����_��� As Integer = 2
    Const cst����_������� As Integer = 3
    Const cst��Ŀ_�˲� As String = "�ɱ���,�ɹ���,�ۼ�,���"
    Const cst��Ŀ_��� As String = "���,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"
    Const cst��Ŀ_������� As String = "�ɹ���,����,�ɱ���,�ɱ����,���,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� From ���ݻ��ڿ��� Where ���� = [1] And ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���ݻ��ڿ���", int����, int����)
    
    If Not rsTmp.EOF Then
        strControlItem = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        
        strControlItem = Replace(strControlItem, "�����", "�ɱ���")
        strControlItem = Replace(strControlItem, "������", "�ɱ����")
    End If
    
    If strControlItem = "" Then
        Select Case int����
            Case cst����_�⹺
                Select Case int����
                    Case cst����_�˲�
                        strControlItem = cst��Ŀ_�˲�
                    Case cst����_���
                        strControlItem = cst��Ŀ_���
                    Case cst����_�������
                        strControlItem = cst��Ŀ_�������
                End Select
        End Select
    End If
    
    GetControlItem = strControlItem
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get����(ByVal int���� As Integer, ByVal int��λ As Integer) As Integer
    '���ܣ��������سɱ��ۺ��ۼۡ�������������ĳ���
    '����1��int����=1 �ɱ���;int����=2 ���ۼ�;int����=3 ����
    '����2��int��λ=1 �ۼ�;int��λ=2 ����;int��λ=3 סԺ;int��λ=4 ҩ��
    '����ֵ�����ݲ����жϾ��ȴ�С
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    On Error GoTo errHandle
    
    gstrSQL = "Select ����,��λ,Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = 0 And ��� = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����")
    
    strFilter = " ����=" & int���� & " And ��λ=" & int��λ
    rsTemp.Filter = strFilter
    
    If rsTemp.RecordCount > 0 Then
        Get���� = rsTemp!����
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'ȡϵͳ����ֵ
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    gtype_UserSysParms.P9_���ý���λ�� = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gtype_UserSysParms.P29_ָ�������۶��۵�λ = Val(zlDatabase.GetPara(29, glngSys, , 0))
    gtype_UserSysParms.P44_����ƥ�� = Val(zlDatabase.GetPara(44, glngSys, , 11))
    gtype_UserSysParms.P54_ʱ��ҩƷ�ԼӼ������ = Val(zlDatabase.GetPara(54, glngSys, , 0))
    gtype_UserSysParms.P64_������� = Val(zlDatabase.GetPara(64, glngSys, , 0))
    gtype_UserSysParms.P75_�⹺�����Ҫ�˲� = Val(zlDatabase.GetPara(75, glngSys, , 0))
    gtype_UserSysParms.P76_ʱ��ҩƷֱ��ȷ���ۼ� = Val(zlDatabase.GetPara(76, glngSys, , 0))
    gtype_UserSysParms.P126_ʱ��ҩƷ�ۼۼӳɷ�ʽ = Val(zlDatabase.GetPara(126, glngSys, , 0))
    gtype_UserSysParms.P149_Ч����ʾ��ʽ = Val(zlDatabase.GetPara(149, glngSys, , 0))
    gtype_UserSysParms.P150_ҩƷ���������㷨 = Val(zlDatabase.GetPara(150, glngSys, , 1))
    gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = Val(zlDatabase.GetPara(173, glngSys, , 0))
    gtype_UserSysParms.P181_ҩƷ��ⰴ�ֶμӳ� = Val(zlDatabase.GetPara(181, glngSys, , 0))
    gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� = Val(zlDatabase.GetPara(183, glngSys, , 0))
    gtype_UserSysParms.P221_ҩƷ���ʱ�� = Val(zlDatabase.GetPara(221, glngSys, , 0))
    gtype_UserSysParms.P275_���۹���ģʽ = Val(zlDatabase.GetPara(275, glngSys, , 0))
    gtype_UserSysParms.P294_����ȡĿ¼�в�����Ϣ = Val(zlDatabase.GetPara(294, glngSys, , 0))
    
    'ȡҩƷ���������
    gstrSQL = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ����")
    gtype_UserDrugDigits.Digit_��� = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_�ɱ��� = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_���ۼ� = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_���� = rs.Fields(3).NumericScale
    
    'ȡҩƷ�ۼ۵�λС��λ��
    gstrSQL = "Select ����, Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ��λ = 1 "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ�ۼ۵�λС��λ��")
    
    If rs.RecordCount > 0 Then
        rs.Filter = "����=1"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_�ɱ��� = rs!����
        
        rs.Filter = "����=2"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_���ۼ� = rs!����
        
        rs.Filter = "����=3"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_���� = rs!����
        
        If gtype_UserSaleDigits.Digit_�ɱ��� < 2 Or gtype_UserSaleDigits.Digit_�ɱ��� > gtype_UserDrugDigits.Digit_�ɱ��� Then
            gtype_UserSaleDigits.Digit_�ɱ��� = gtype_UserDrugDigits.Digit_�ɱ���
        End If
        
        If gtype_UserSaleDigits.Digit_���ۼ� < 2 Or gtype_UserSaleDigits.Digit_���ۼ� > gtype_UserDrugDigits.Digit_���ۼ� Then
            gtype_UserSaleDigits.Digit_���ۼ� = gtype_UserDrugDigits.Digit_���ۼ�
        End If
        
        If gtype_UserSaleDigits.Digit_���� < 2 Or gtype_UserSaleDigits.Digit_���� > gtype_UserDrugDigits.Digit_���� Then
            gtype_UserSaleDigits.Digit_���� = gtype_UserDrugDigits.Digit_����
        End If
    End If
    
    'ҩƷ������ʾ��ʽ
    gintҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ", , , 2))
    gint����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ"))
    
    If gintҩƷ������ʾ < 0 Or gintҩƷ������ʾ > 2 Then gintҩƷ������ʾ = 2
    If gint����ҩƷ��ʾ < 0 Or gint����ҩƷ��ʾ > 1 Then gint����ҩƷ��ʾ = 0
    
    '���뷽ʽ
    gint���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ"))
    If gint���뷽ʽ < 0 Or gint���뷽ʽ > 1 Then gint���뷽ʽ = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'����ָ���ָⷿ�����÷�Χ�ĵ�λ
Public Function GetSpecUnit(ByVal lng�ⷿID As Long, ByVal int��Χ As Integer) As String
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(����,1) AS ��λ From ҩƷ�ⷿ��λ Where �ⷿID=[1] And ���÷�Χ=[2]"
    Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ", lng�ⷿID, int��Χ)

    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!��λ
    Else
'        MsgBox "�ÿⷿδ���ÿⷿ��λ�����ݲ��������Լ��������ȡȱʡ��λ��" & _
'            vbCrLf & "ȱʡ��λ�Ĺ���" & _
'            vbCrLf & "  ���������סԺ�������סԺ�ģ�ȡסԺ��λ" & _
'            vbCrLf & "  ������������ģ�ȡ���ﵥλ" & _
'            vbCrLf & "  ����ҩ�����Եģ�ȡҩ�ⵥλ" & _
'            vbCrLf & "  ����ȡ�ۼ۵�λ", vbInformation, gstrSysName
        
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿID)

        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            'סԺ��λ
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '���ﵥλ
            strUnit = 2
        ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            strUnit = 4
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            strUnit = 1
        End If
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    GetSpecUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetSpecUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'ȡҩƷ��λ����
Public Function GetDrugUnit(ByVal lng�ⷿID As Long, ByVal frmCaption As String) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim intUnit As Integer, strUnit As String
    Dim blnȱʡ As Boolean
    Dim lngModul As Long
    On Error GoTo ErrHand
    
    If frmCaption Like "ҩƷ�⹺������*" Then
        lngModul = 1300
    ElseIf frmCaption Like "ҩƷ����������*" Then
        lngModul = 1301
    ElseIf frmCaption Like "ҩƷ����������*" Then
        lngModul = 1302
    ElseIf frmCaption Like "����۵�������*" Then
        lngModul = 1303
    ElseIf frmCaption Like "ҩƷ�ƿ����*" Then
        lngModul = 1304
    ElseIf frmCaption Like "ҩƷ���ù���*" Then
        lngModul = 1305
    ElseIf frmCaption Like "ҩƷ�����������*" Then
        lngModul = 1306
    ElseIf frmCaption Like "ҩƷ�̵����*" Then
        lngModul = 1307
    ElseIf frmCaption Like "ҩƷ��ۼ���*" Then
        lngModul = 1308
    ElseIf frmCaption Like "ҩƷ�ƻ�����*" Or frmCaption Like "ҩƷ�ɹ��ƻ�*" Then
        lngModul = 1330
    ElseIf frmCaption Like "ҩƷ��������*" Then
        lngModul = 1331
    ElseIf frmCaption Like "ҩƷ�������*" Then
        lngModul = 1343
    End If
    
    intUnit = 0
    '��������쵥����ֱ�ӷ���ע����еĵ�λ
    If lngModul > 0 And lngModul <> 1331 And lngModul <> 1307 And lngModul <> 1308 Then
        intUnit = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, lngModul))
        '���ز������õĵ�λ˳�����£�0-ȱʡ;1-ҩ��;2-����;3-סԺ;4-�ۼۣ���Ҫת��Ϊ��ϵͳ������һ��
        If intUnit = 1 Then
            intUnit = 4
        ElseIf intUnit = 4 Then
            intUnit = 1
        End If
        strUnit = intUnit
    End If
    
    If intUnit = 0 Then
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿID)

        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            intUnit = 1
            strUnit = 4
        ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '���ﵥλ
            intUnit = 2
            strUnit = 2
        ElseIf InStr(strobjTemp, "2") <> 0 Then
            'סԺ��λ
            intUnit = 3
            strUnit = 3
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            intUnit = 4
            strUnit = 1
        End If
        
        'ȡ��ҩ��ȱʡ��ʹ�õĵ�λ
        GetDrugUnit = GetSpecUnit(lng�ⷿID, intUnit)
    Else
        GetDrugUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetDrugUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "�ۼ۵�λ"
End Function

Public Function MediWork_GetCheckStockRule(ByVal lng�ⷿID As Long) As Integer
    'ȡ���������
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���������", lng�ⷿID)

    If Not rsData.EOF Then
        MediWork_GetCheckStockRule = rsData!�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ɱ���(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ����=1 and ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID, lng�ⷿID, lng����)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get�ɱ��� = rsData!ƽ���ɱ���
    Else
        '����޷��ӿ����ȡ�ɱ��ۣ����ҩƷ�����ȡ
        gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID)
        If Not rsData.EOF Then
            If Val(NVL(rsData!�ɱ���, 0)) > 0 Then
                Get�ɱ��� = rsData!�ɱ���
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function Get���ۼ�(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    '���ܣ���ȡʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����)
    
    If rsData.EOF Then
        'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
        '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
        dblָ�����ۼ� = rsData!ָ�����ۼ�
        dbl��������� = rsData!���������
        
        Get���ۼ� = 0
        dbl�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng����)
        dbl�ӳ��� = rsData!�ӳ��� / 100
        dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
        dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
        Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
    Else
        If rsData!���ۼ� = 0 Then
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get���ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            Get���ۼ� = rsData!���ۼ� * dbl����ϵ��
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�����룬���ƣ���������ĳһ��
Public Function FindRow(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindRow = True
    With mshBill
        If .rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .MsfObj.TopRow = .Row
                    .SetRowColor CLng(intRow), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.���� " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.�շ�ϸĿid " & _
                  "    FROM �շ���Ŀ���� A" & _
                  "    Where A.���� LIKE [1]) a," & _
                  " �շ���ĿĿ¼ B " & _
                  " Where a.�շ�ϸĿid = b.ID"
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, "����ָ��ҩƷ", IIf(gstrMatchMethod = "0", "%", "") & str�Ƚ�ֵ & "%")
        
        If rsCode.EOF Then
            FindRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .MsfObj.TopRow = .Row
                        .SetRowColor CLng(intRow), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindRow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'���ܣ������������ŷ������ܷ��ʵĲ���
'���أ���'����ҩ'��'�г�ҩ',�ձ�ʾ����
Public Function GetMaterial(lngUnitID As Long) As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(gstrprivs, "����ҩ��") > 0 Then Exit Function
    
    On Error GoTo errH
    
    rsTmp.CursorLocation = adUseClient

    gstrSQL = "Select A.��������,B.���� From ��������˵�� A,���ű� B Where A.����ID=B.ID And B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡָ�����ŵĹ�������", lngUnitID)
    
    If rsTmp.EOF Then Exit Function
    
    rsTmp.Filter = "��������='��ҩ��' or ��������='��ҩ��' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'����ҩ'"
    
    rsTmp.Filter = "��������='��ҩ��' or ��������='��ҩ��' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'�г�ҩ'"
    
    rsTmp.Filter = "��������='��ҩ��' or ��������='��ҩ��' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'�в�ҩ'"
    
    GetMaterial = Mid(GetMaterial, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Public Function ExecuteSql(ByRef arrSql As Variant, strTitle As String, _
Optional ByVal blnCommit As Boolean = True, Optional ByVal blnBeginTrans As Boolean = True) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer
    Dim intouter As Integer
    Dim intInner As Integer
    
    ExecuteSql = False
    If UBound(arrSql) >= 0 Then
        '��SQL���а�ҩƷID��������
        intouter = UBound(arrSql) - 1
        If Split(arrSql(UBound(arrSql)), ":")(0) = "����" Then
            intouter = UBound(arrSql) - 2
        Else
            intouter = UBound(arrSql) - 1
        End If
        
        For i = 0 To intouter
            For j = i + 1 To intouter + 1
                If CLng(Split(arrSql(j), ";")(0)) < CLng(Split(arrSql(i), ";")(0)) Then
                    strTmp = CStr(arrSql(j))
                    arrSql(j) = arrSql(i)
                    arrSql(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        If blnBeginTrans Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(Split(arrSql(i), ";")(1)), strTitle)
        Next
        If blnCommit Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
       
errH:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'ȡָ����ͷ����λ��
Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    Dim i As Integer
    
    GetCol = -1
    
    If TypeName(mshFlex) = "MSHFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    ElseIf TypeName(mshFlex) = "VSFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    End If
End Function

'����ҩƷ������Ʊ�����ݣ���ȡ�Է��ⷿ
'Writed by zyb
'-----------------����-----------------
'���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (1"������Է��ⷿ",3"��˫����ͨ")
'�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (2"���������ڿⷿ",3"��˫����ͨ")
'-----------------����-----------------
'���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (2"���������ڿⷿ",3"��˫����ͨ")
'�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (1"������Է��ⷿ",3"��˫����ͨ")
Public Function ReturnSQL(ByVal lng�ⷿID As Long, ByVal strCaption As String, _
    Optional ByVal bln���� As Boolean = True, _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    
    Dim str�ⷿ���� As String, strҩƷ���� As String, strվ������ As String, strSQL As String
    
    On Error GoTo errHandle
    strվ������ = GetDeptStationNode(lng�ⷿID)
    str�ⷿ���� = "('H','I','J','K','L','M','N')"
    
    strҩƷ���� = ",(Select �Է��ⷿID ID From ҩƷ�������" & _
            " Where ���ڿⷿID=[1] And ���� In (" & IIf(bln����, 1, 2) & ",3)" & _
            " Union" & _
            " Select ���ڿⷿID ID From ҩƷ�������" & _
            " Where �Է��ⷿID=[1] And ���� In (" & IIf(bln����, 2, 1) & ",3)) D"
    Select Case lngModuleNO
        Case 1304   'ҩƷ�ƿ����
            strSQL = " SELECT DISTINCT a.id,a.����,a.����" & _
                    " FROM ��������˵�� c, �������ʷ��� b, ���ű� a" & strҩƷ���� & _
                    " Where c.�������� = b.����" & _
                    "   AND b.����||'' in " & str�ⷿ���� & _
                    "   AND a.id = c.����id And A.ID=D.ID " & _
                    "   AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.����"
        Case Else
            strSQL = " SELECT DISTINCT a.id,a.����,a.����" & _
                    " FROM ��������˵�� c, �������ʷ��� b, ���ű� a" & strҩƷ���� & _
                    " Where c.�������� = b.����" & _
                    "   AND b.����||'' in " & str�ⷿ���� & _
                    "   AND a.id = c.����id And A.ID=D.ID" & IIf(strվ������ <> "", " AND (a.վ��=[2] or a.վ�� is null) ", "") & _
                    "   AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.����"
    End Select
    
    Set ReturnSQL = zlDatabase.OpenSQLRecord(strSQL, strCaption, lng�ⷿID, strվ������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ��ͬ����(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_���� As Boolean, blnSecond_���� As Boolean
    
    ��ͬ���� = False
    
    If sinFirst = 0 Or sinSecond = 0 Then '0��������֮��
        ��ͬ���� = True
        Exit Function
    End If
    
    blnFirst_���� = (sinFirst <= 0)
    blnSecond_���� = (sinSecond <= 0)
    
    ��ͬ���� = (blnFirst_���� = blnSecond_����)
End Function

'��ָ���п�ʼ�������
Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng����� As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    
    With mshBill
        lngRows = .rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng�����) = lngRow
        Next
    End With
End Sub

'ת����ֵΪ����
Public Function TranNumToDate(ByVal strNum As String) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim StrDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    TranNumToDate = StrDate
End Function

'��ȡָ������ĸ�����
Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function

'��ȡָ������ı���
Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    
    On Error Resume Next
   
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlStr.TruncZero(strCaption)
End Function

Public Sub CheckLapse(ByVal strЧ�� As String)
    'ʧЧҩƷ���
    If Not IsDate(strЧ��) Then Exit Sub
    
    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
        '����ΪʧЧ��
        strЧ�� = Format(DateAdd("D", 1, CDate(strЧ��)), "yyyy-mm-dd")
    End If
    
    If Format(strЧ��, "yyyy-MM-dd") < Format(Sys.Currentdate, "yyyy-MM-dd") Then
        MsgBox "��ҩƷ�Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
    End If
End Sub

'ҩƷ�������ʱ���Ƿ��ж�������������ˣ��䷵����˽��
Public Function ҩƷ�������(ByVal str������ As String) As Boolean
    Dim blnBillVerify As Boolean
    
    ҩƷ������� = True
    
    blnBillVerify = IIf(gtype_UserSysParms.P64_������� = 0, False, True)
    If Not blnBillVerify Then Exit Function
    
    ҩƷ������� = (Trim(str������) <> Trim(UserInfo.�û�����))
    If Not ҩƷ������� Then MsgBox "������������˲�����ͬһ�ˣ����飡", vbInformation, gstrSysName
End Function
'ͨ��ҩƷѡ��������ҩƷʱ�����ҩƷ����е�������Ӳ������ʡ�ҩƷĿ¼�еķ��������жϳ��Ĳ�һ�£��򱨴�
Public Function ���������(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln����Ƿ���� As Boolean, bln���� As Boolean, bln�ⷿ As Boolean
    
    ��������� = False
    On Error GoTo errHandle
    '���û�п���¼����ֱ���˳�
    gstrSQL = " Select Count(*) ��¼�� From ҩƷ��� " & _
              " Where �ⷿID=[1] And ����=1 And ҩƷID=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���ڿ������", lng�ⷿID, lngҩƷID)
    
    If rsCheck!��¼�� = 0 Then
        ��������� = True
        Exit Function
    End If
    
    '���ڷ�����¼���������
    gstrSQL = " Select Count(*) ���� From ҩƷ��� " & _
              " Where �ⷿID=[1] And ����=1 And Nvl(����,0)<>0 And ҩƷID=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "���������", lng�ⷿID, lngҩƷID)
              
    bln����Ƿ���� = (rsCheck!���� <> 0)
    
    '���ж��Ƿ��ǿⷿ
    gstrSQL = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿID)

    bln�ⷿ = (rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    gstrSQL = " Select Nvl(ҩ�����,0) ��������,nvl(ҩ������,0) ҩ���������� " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷĿ¼�еķ�������", lngҩƷID)
              
    If bln�ⷿ Then
        bln���� = (rsCheck!�������� = 1)
    Else
        bln���� = (rsCheck!ҩ���������� = 1)
    End If
    
    ��������� = (bln����Ƿ���� = bln����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'���ҩƷ�ļ۸��Ƿ�Ϊ���µļ۸񣨰�ҩ�ⵥλ���бȽϣ�ʱ�۲�����ҩƷ����飩�������������
'�����ڱ���ǰ�жϺ��鷳���Ҹ��ֵ��ݵı���б�������ݲ�һ������ˣ����������֮�����ύǰ���ѱ�������ݽ��м��
'ҩƷ��ͬ�ļ�¼�Թ�
Public Function ��鵥��(ByVal lng���� As Long, ByVal strNo As String, Optional ByVal blnMsg As Boolean = True, Optional ByVal bln�ƿⵥ As Boolean = False) As Boolean
    Dim rsPrice As New ADODB.Recordset
    Dim lngҩƷ_Last As Long, lngҩƷ_Cur As Long
    Dim intPriceDigit As Integer
    Dim intCostDigit As Integer
             
    On Error GoTo errHandle
    '�Զ�������鲢ִ�е���
    Call AutoAdjustPrice_ByNO(lng����, strNo)
    
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
    
    '����ҩƷ���շѼ�Ŀȡ���¼۸�ʱ�۷���ҩƷ�ӿ���ȡ���¼۸�ʱ��ҩƷ�����ǰ��������ģ�����޿�����ʾ�޵��ۣ������ϸ���ƿ����������޿��Ҳ����������⣩
        
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid , 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = [1] And a.No = [2] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(b.�ּ�, " & intPriceDigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid , 0 ԭ��, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C ," & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = [1] And a.No = [2] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPriceDigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
                  " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid , 0 ԭ��, decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B ," & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = [1] And a.No = [2] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�)," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ�۸�", lng����, strNo)
    
    If rsPrice.EOF Then
        ��鵥�� = True
        Exit Function
    End If
    
    lngҩƷ_Last = 0
    With rsPrice
        Do While Not .EOF
            lngҩƷ_Cur = !ҩƷID
            If lngҩƷ_Cur <> lngҩƷ_Last Then
                If blnMsg Then
                    If MsgBox("��" & IIf(bln�ƿⵥ, Round(!��� / 2 + 0.49), !���) & "��ҩƷ��" & !���� & "�������¼۸��Ƿ�������浥�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lngҩƷ_Last = lngҩƷ_Cur
            .MoveNext
        Loop
        ��鵥�� = True
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'------------------------------------------------
'���ܣ� ����ת������
'������
'   strOld��ԭ����
'���أ� �������ɵ�����
'------------------------------------------------
Public Function TranPasswd(strOld As String) As String
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function GetBillInfo(ByVal lng���� As Long, ByVal strNo As String, Optional ByVal bln�������� As Boolean = True) As String
    Dim rsBillInfo As New ADODB.Recordset
    
    On Error GoTo errHandle
    '��ȡ���ݵ�����޸�ʱ��
    gstrSQL = " Select to_char(Max(" & IIf(bln��������, "��������", "�������") & "),'yyyyMMddhh24miss') ���� From ҩƷ�շ���¼ " & _
            " Where ����=[1] And NO=[2]"
    Set rsBillInfo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ݵ�����޸�ʱ��", lng����, strNo)
    
    With rsBillInfo
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then Exit Function
        If IsNull(!����) Then Exit Function
        GetBillInfo = !����
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'ȡҩƷ���۸��������С��λ��
Public Function GetDigit(ByVal int���� As Integer, ByVal int��� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���ʣ�0-���㾫��;
    'int���1-ҩƷ;2-����
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ҩƷ��λ:1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If int���� = 4 Then 'ȡ��� ��λ=5�Ĳ��ǽ��
        int��λ = 5
    End If
    
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = [1] And ��� = [2] And ���� = [3] And ��λ = [4] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int����, int���, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!����
    End If
    
    If GetDigit = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        GetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
End Function

'���ݿⷿ�İ�װ��λ��ȡҩƷ�ļ۸����������С��λ�������㾫�ȣ�
Public Sub GetDrugDigit(ByRef lng�ⷿID As Long, ByVal frmCaption As String, ByRef intUnit As Integer, ByRef intCostDigit As Integer, ByRef intPriceDigit As Integer, ByRef intNumberDigit As Integer, ByRef intMoneyDigit As Integer)
    Dim strUnit As String
    Dim intTemp As Integer
    
    Const conInt���� As Integer = 0
    
    Const conIntҩƷ As Integer = 1
    
    Const conint�ۼ۵�λ As Integer = 1
    Const conint���ﵥλ As Integer = 2
    Const conintסԺ��λ As Integer = 3
    Const conintҩ�ⵥλ As Integer = 4
        
    Const conInt�ɱ��� As Integer = 1
    Const conInt�ۼ� As Integer = 2
    Const conInt���� As Integer = 3
    Const conInt��� As Integer = 4
    
    If lng�ⷿID > 0 Then
        If frmCaption Like "ҩƷ���չ���*" Then
            strUnit = conintҩ�ⵥλ
        Else
            strUnit = GetDrugUnit(lng�ⷿID, frmCaption)
        
            Select Case strUnit
                Case "�ۼ۵�λ"             '�ۼ۵�λ����Ҫ���Ƽ���
                    intUnit = conint�ۼ۵�λ
                Case "���ﵥλ"
                    intUnit = conint���ﵥλ
                Case "סԺ��λ"
                    intUnit = conintסԺ��λ
                Case "ҩ�ⵥλ"
                    intUnit = conintҩ�ⵥλ
            End Select
        End If
    Else
        
        If frmCaption Like "ҩƷ�ƻ�����*" Or frmCaption Like "ҩƷ�ɹ��ƻ�*" Then
            intTemp = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, 1330))
            Select Case intTemp
            Case 1 'ҩ��
                intUnit = conintҩ�ⵥλ
            Case 2  '����
                intUnit = conint���ﵥλ
            Case 3  'סԺ
                intUnit = conintסԺ��λ
            Case 4  '�ۼ�
                intUnit = conint�ۼ۵�λ
            Case Else
                intUnit = conintҩ�ⵥλ
            End Select
        Else
            intUnit = conintҩ�ⵥλ
        End If
    End If

    '�ֱ�ȡҩƷ�ɱ��ۡ��ۼۡ�����������С��λ��
    intCostDigit = GetDigit(conInt����, conIntҩƷ, conInt�ɱ���, intUnit)
    intPriceDigit = GetDigit(conInt����, conIntҩƷ, conInt�ۼ�, intUnit)
    intNumberDigit = GetDigit(conInt����, conIntҩƷ, conInt����, intUnit)
    intMoneyDigit = GetDigit(conInt����, conIntҩƷ, conInt���)

End Sub

Public Function Select����ѡ����(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional strSQL As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim strPa As String
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    strPa = zlDatabase.GetPara(44, glngSys, 0): strPa = IIf(strPa = "", "11", strPa)
    
    If strSQL <> "" Then
    
        gstrSQL = strSQL
    Else
        gstrSQL = "" & _
        "   Select distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c" & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.���� in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) ") & _
            "         AND a.id = c.����id " & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[4] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(strPa, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlStr.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(strPa, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlStr.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSQL = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSQL = "" Then
        '�����¼�
        Set rsTemp = zlDatabase.ShowSQLSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.�û�ID, str��������, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "û�����������Ĳ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        objCtl.Tag = Val(rsTemp!id)
    End If
    zlCommFun.PressKey vbKeyTab
    Select����ѡ���� = True
End Function

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub CostPrice()
    '�Ƿ�����ҩ����Ա�鿴���ݵĳɱ���
    mblnCostPrice = IIf(gtype_UserSysParms.P85_ҩ���鿴���ݳɱ��� = 1, True, False)
End Sub

Public Function DepotProperty(ByVal lng��Աid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHandle
    '����ָ����Ա�Ƿ����ҩ������
    gstrSQL = "Select Distinct �������� From ������Ա B,��������˵�� A " & _
             " Where A.�������� like '%ҩ��' And " & _
             " A.����id = B.����id And B.��Աid = [1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng��Աid)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ShowCostPrice() As Boolean
    'ҩ����Ա���ܣ�ֻ��ҩ����Ա���Բ�������Ϊ׼
    Call CostPrice
    If DepotProperty(UserInfo.�û�ID) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = mblnCostPrice
    End If
End Function

Public Function CheckNOExists(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where NO=[1] And ����=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���ڸõ���", strNo, int����)
    
    If rsTemp.RecordCount = 0 Then Exit Function
    CheckNOExists = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�жϸ�ҩƷ�ڵ�ǰ���Ŀ���Ƿ���ڿ�����ޣ����򷵻���
Public Function IsLowerLimit(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long) As Boolean
    Dim dbl������� As Double, dbl���� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '��ȡ�������
    gstrSQL = " Select Sum(Nvl(ʵ������,0)) AS ������� From ҩƷ���" & _
              " Where ����=1 And �ⷿID=[1] And ҩƷID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ���ⷿ��ʵ�ʿ��", lng�ⷿID, lngҩƷID)
    
    If rsTemp.RecordCount = 1 Then dbl������� = NVL(rsTemp!�������, 0)
    
    '��ȡ�����޶��е�����
    gstrSQL = " Select Nvl(����,0) AS ���� From ҩƷ�����޶�" & _
              " Where �ⷿID=[1] And ҩƷID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����޶��е�����", lng�ⷿID, lngҩƷID)
    
    If rsTemp.RecordCount = 1 Then dbl���� = rsTemp!����
    
    IsLowerLimit = (dbl������� < dbl����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'���뷽ʽ
'staVal: StartusBar�ؼ�
'bytType: 0=ƴ��; 1=���;  ��ǰ����״̬
    Dim i As Integer
    For i = 1 To staVal.Panels.count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "���뷽ʽ", 0
                gint���뷽ʽ = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "���뷽ʽ", 1
                gint���뷽ʽ = 1
            End If
        End If
    Next
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'��ȡ��������վ����Ϣ
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo errHandle
    strTmp = "select վ�� from ���ű� where id=[1]"
    Set rsSQL = zlDatabase.OpenSQLRecord(strTmp, "��ȡ��������վ����Ϣ", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = NVL(rsSQL!վ��)
    End If
    rsSQL.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVSFlexRows(ByVal vsfVal As VSFlexGrid, Optional ByVal blnHidden = False) As Long
'--------------------------------------------------------------
'���ܣ���VSFlexGrid��������������ͷ��
'������
'  blnHidden��True��������ص�������False�������ص�������
'���أ�������
'--------------------------------------------------------------
    Dim i As Long, lngRows As Long
    For i = 0 To vsfVal.rows - 1
        If blnHidden Then
            If vsfVal.RowHidden(i) Then lngRows = lngRows + 1
        Else
            If vsfVal.RowHidden(i) = False Then lngRows = lngRows + 1
        End If
    Next
    GetVSFlexRows = lngRows
End Function

Public Sub SetSelectorRS( _
    ByVal byt�༭ģʽ As Byte, _
    ByVal strModeName As String, _
    Optional ByVal lng��Դ�ⷿ As Long = 0, _
    Optional ByVal lngĿ��ⷿ As Long = 0, _
    Optional ByVal lngʹ�ò��� As Long = 0, _
    Optional ByVal lng��Ӧ�� As Long = 0, _
    Optional ByVal byt���÷�ʽ As Byte = 0, _
    Optional ByVal bln����ͣ��ҩƷ As Boolean = False, _
    Optional ByVal bln���޴洢�ⷿҩƷ As Boolean = False, _
    Optional ByVal byt�̵㵥�� As Byte = 0, _
    Optional ByVal bln����� As Boolean = True, _
    Optional ByVal bln���� As Boolean = False, _
    Optional ByVal bln���Է������ As Boolean = True, _
    Optional ByVal str�̵�ʱ�� As String = "" _
    )
'----------------------------------------------------------------------------------------
'���ܣ���ʼ��grsMaster��grsMasterInput��grsSlave����
'      Ϊ����ҩƷѡ����(frmSelector)������׼����
'������
'  byt�༭ģʽ�� 1����⣻ 2������
'  lng��Դ�ⷿ��
'----------------------------------------------------------------------------------------
    Const CON_FMT = "'999999999990.99999'"
    
    Dim strSQL As String, strTmp As String
    Dim strUnit As String, strConversionUnit As String
    Dim rsTemp As ADODB.Recordset
    Dim IntStockCheck As Integer
    Dim intUnit As Integer, intCostDigit As Integer, intPriceDigit As Integer, intNumberDigit As Integer, intMoneyDigit As Integer
    Dim str�̵�sql As String
    
    On Error GoTo errHandle
    With grsMaster
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsMasterInput
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsSlave
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    
    '������λ
    If strModeName = "ҩƷ�������" Or strModeName = "ҩƷ�ƿ����" Then
        Call GetDrugDigit(lngʹ�ò���, strModeName, intUnit, intCostDigit, intPriceDigit, intNumberDigit, intMoneyDigit)
    Else
        Call GetDrugDigit(IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ), strModeName, intUnit, intCostDigit, intPriceDigit, intNumberDigit, intMoneyDigit)
    End If
    Select Case intUnit
        Case 1: strConversionUnit = "1"
        Case 2: strConversionUnit = "d.�����װ"
        Case 3: strConversionUnit = "d.סԺ��װ"
        Case Else
            strConversionUnit = "d.ҩ���װ"
    End Select
    
    '�����
    If bln����� = True And (strModeName = "ҩƷ�������" Or strModeName = "ҩƷ���ù���" Or strModeName = "ҩƷ�ƿ����") Then
        If strModeName = "ҩƷ�������" Then bln����� = (Val(zlDatabase.GetPara("ҩƷ�����γ���", glngSys, 1343, 0)) = 1)
        If strModeName = "ҩƷ���ù���" Then bln����� = (Val(zlDatabase.GetPara("ҩƷ�����γ���", glngSys, 1305, 0)) = 1)
        If strModeName = "ҩƷ�ƿ����" Then bln����� = (Val(zlDatabase.GetPara("ҩƷ�����γ���", glngSys, 1304, 0)) = 1)
    End If
    
    '��鲢ִ�е���
    Call AutoAdjustPrice_Batch
    
    '��ȡ����������ȷ����治��Ĳ���ȡ����
    strSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ƿ���������", lng��Դ�ⷿ)
    If Not rsTemp.EOF Then IntStockCheck = NVL(rsTemp!�����, 0)
    rsTemp.Close
    
    '*ѡ��ģʽ�����ݼ�*'
    strSQL = _
        "Select " & _
        " d.����,d.��ҩ��̬, d.ҩ������, d.ͨ������, d.ҩƷ��Դ As ��Դ, d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, d.ҩƷ����, " & _
        " d.��Ʒ��, d.���, d.���� As ������, Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����) as ԭ����, d.ҩ��id, d.ҩƷid, " & _
        " trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) �ϴβɹ���, " & _
        " trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, p.�ۼ�, s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPriceDigit, "0") & "')) �ۼ�, " & _
        " d.�ۼ۵�λ, d.����ϵ�� As �ۼ۰�װ," & _
        " d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ, d.ҩ���װ, " & _
        " trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')) ��������, " & _
        " s.�������, s.�����, s.�����,  d.���Ч�� ��Ч��, d.ҩ�����, d.ҩ������, d.ʱ��," & _
        " trim(to_char(d.ָ�������� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ��������, " & _
        " trim(to_char(d.ָ�����ۼ� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ�����ۼ�, " & _
        " d.�ӳ���, e.�ⷿ��λ, d.��׼�ĺ�, s.������� ʵ������, " & _
        " s.��������, d.��ͬ��λ, d.ҩ�ۼ���,e.���ñ�־,d.ͣ��,d.�ϴι�Ӧ�� " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.���� ����,Decode(c.���, '7', Decode(d.��ҩ��̬, 1, '��Ƭ', 2, '����', 'ɢװ'), '') As ��ҩ��̬,A.���� ��Ʒ��, C.���� ҩ������,C.���� ͨ������, 0 AS ҩ��ID,C.���� ҩƷ����,C.���� ҩƷ����," & vbNewLine & _
        "     C.���,C.����,d.ԭ����,C.���,C.���㵥λ AS �ۼ۵�λ,DECODE(C.�Ƿ���,1,'��','��') ʱ��,D.ҩƷ��Դ,D.����ҩ��,D.��׼�ĺ�, D.ҩ��ID," & vbNewLine & _
        "     D.ҩƷID, nvl(to_char(D.���Ч��,'9999990'),0) ���Ч��," & vbNewLine & _
        "     DECODE(D.ҩ�����,1,'��','��') ҩ�����,DECODE(D.ҩ������,1,'��','��') ҩ������," & vbNewLine & _
        "     to_char(D.����ϵ��, " & CON_FMT & ") ����ϵ��," & vbLf & _
        "     D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & vbNewLine & _
        "     D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & vbNewLine & _
        "     D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & vbNewLine & _
        "     D.ָ��������,d.ָ�����ۼ�, nvl(D.�ɱ���,0) ��ʼ�ɱ���,D.�ӳ���,D.ҩ�ۼ���," & vbNewLine & _
        "     M.����ID AS ��;����ID,M.���㵥λ AS ������λ,Q.���� As ��ͬ��λ,Decode(Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '��','��') As ͣ��,f.���� �ϴι�Ӧ�� " & vbNewLine
    strSQL = strSQL & _
        "   FROM �շ���ĿĿ¼ C,ҩƷ��� D,�շ���Ŀ���� A,ҩƷ���� J,ҩƷ���� T,������ĿĿ¼ M,��Ӧ�� Q, ���Ʒ���Ŀ¼ E, ��Ӧ�� F " & vbNewLine
        
    If bln���� = False Then
        strSQL = strSQL & IIf(lng��Դ�ⷿ <> 0, " ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K", "") & vbNewLine & _
        IIf(lngĿ��ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[3] Group By ִ�п���ID,�շ�ϸĿID) I ", "") & vbNewLine
    End If
    strSQL = strSQL & "   WHERE C.ID=D.ҩƷID AND D.ҩ��ID=T.ҩ��ID AND T.ҩ��ID=M.ID and m.����id=e.id AND M.��� IN ('5','6','7') and t.�ٴ��Թ�ҩ is null And d.�ϴι�Ӧ��id = f.id(+) "
    
    If bln���� = False Then
        strSQL = strSQL & IIf(lng��Դ�ⷿ <> 0, "     And D.ҩƷID=K.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        IIf(lngĿ��ⷿ <> 0, "     And D.ҩƷID=I.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "")
    End If
    
    If bln���Է������ = False Then
        strSQL = strSQL & " And" & _
             " (Decode(c.�������, 1, 1, 3, 1, 0) = " & _
             " (Select Distinct '1' From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2] And ������� In (1, 3)) Or " & _
             " Decode(c.�������, 2, 1, 3, 1, 0) =" & _
             " (Select Distinct '1' From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2] And ������� In (2, 3)) Or Exists" & _
             " (Select 1 From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2])) "
    End If
    
    strSQL = strSQL & _
        "     AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
        "     And (C.վ�� = [1] or c.վ�� is null) AND T.ҩƷ����=J.����(+) And D.��ͬ��λID=Q.ID(+) " & _
        IIf(bln����ͣ��ҩƷ = False, " And (C.����ʱ�� Is Null Or To_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "(Select �շ�ϸĿid, �ּ� �ۼ� " & _
        " From �շѼ�Ŀ Where (Sysdate Between ִ������ And ��ֹ���� or Sysdate>=ִ������ And ��ֹ���� Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine
    If byt���÷�ʽ = 1 Then
       '��������ҩ
       strSQL = strSQL & _
           "(Select a.ҩƷid,Max(�ϴβ���) AS ����,max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(nvl(ʵ������,0)), 0, null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�," & _
           " To_Char(Sum(b.ʵ������), '99999999999990.99') �������� " & vbNewLine & _
           "From ҩƷ��� A, ҩƷ���� B " & vbNewLine & _
           "Where a.����=1 and a.ҩƷid=b.ҩƷid And a.�ⷿid=b.�ⷿid and b.����id=[3] and b.�ڼ�=to_date(sysdate,'yyyy') "
    Else
       '��ҩ����ҩ
       strSQL = strSQL & _
           "(Select a.ҩƷid, Max(a.�ϴβ���) AS ����, max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " Sum(a.ʵ������) �������," & _
           " Sum(a.ʵ�ʽ��) �����," & _
           " Sum(a.ʵ�ʲ��) �����," & _
           " Decode(Sum(nvl(ʵ������,0)), 0, null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�," & _
           " '' �������� " & vbNewLine & _
           "From ҩƷ��� A " & vbNewLine & _
           "Where ����=1 "
    End If
    If lng��Դ�ⷿ <> 0 Or lngĿ��ⷿ <> 0 Then
       strSQL = strSQL & " And a.�ⷿID=" & IIf(lng��Դ�ⷿ = 0, "[3]", "[2]")
    End If
    strSQL = strSQL & vbNewLine & _
       "Group By a.ҩƷid) S," & vbNewLine & _
       "(Select ҩƷID,�ⷿID,�ⷿ��λ,���ñ�־ From ҩƷ�����޶� Where �ⷿID=[2]) E " & vbNewLine & _
       "Where D.ҩƷID=P.�շ�ϸĿID And D.ҩƷID=S.ҩƷID" & IIf(Not (IntStockCheck = 2 And byt�༭ģʽ = 2) Or byt�̵㵥�� = 1 Or Not bln�����, "(+)", "") & _
       "  And D.ҩƷID=E.ҩƷID(+) " & vbNewLine & _
       "Order By D.ҩ������,D.ҩƷ���� "
    Set grsMaster = zlDatabase.OpenSQLRecord(strSQL, "ҩƷ���", gstrNodeNo, lng��Դ�ⷿ, lngĿ��ⷿ)
    
    
    '*¼��ģʽ�����ݼ�*'
    strSQL = _
        "Select " & _
        " d.����,d.ҩ������, d.ͨ������, d.ҩƷ��Դ ��Դ, d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, f.���� ҩƷ����, " & _
        " d.��Ʒ��, d.���, d.���� As ������, Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����) as ԭ����, d.ҩ��id, d.ҩƷid, " & _
        " trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) �ϴβɹ���, " & _
        " trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, Nvl(d.�ϴ��ۼ�,p.�ۼ�), s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPriceDigit, "0") & "')) �ۼ�, " & _
        " d.�ۼ۵�λ, d.����ϵ�� �ۼ۰�װ, " & _
        " d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ, d.ҩ���װ, " & _
        " trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')) ��������, " & _
        " s.�������,s.�����, s.�����, d.���Ч�� ��Ч��, d.ҩ�����, d.ҩ������, d.ʱ��, " & _
        " trim(to_char(d.ָ��������* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ��������, " & _
        " trim(to_char(d.ָ�����ۼ�* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ�����ۼ�, " & _
        " d.�ӳ���, e.�ⷿ��λ, d.��׼�ĺ�, s.������� ʵ������," & _
        " s.��������, d.��ͬ��λ, d.ҩ�ۼ���,e.���ñ�־, Max(Decode(f.����, '1', f.����, Null)) ����, Max(Decode(f.����, '3', f.����, Null)) ���ּ���, Max(Decode(f.����, '2', f.����, Null)) �����,d.ͣ��,d.�ϴι�Ӧ�� " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.���� ����,Decode(c.���, '7', Decode(d.��ҩ��̬, 1, '��Ƭ', 2, '����', 'ɢװ'), '') As ��ҩ��̬,C.���� ҩ������,C.���� AS ͨ������,0 AS ҩ��ID,M.����ID AS ��;����ID,M.���㵥λ AS ������λ, " & _
        "   C.���� AS ҩƷ����, a.���� As ��Ʒ��, c.���, c.����, d.ԭ����, d.ҩƷ��Դ, d.����ҩ��, d.��׼�ĺ�, d.ҩ��id, " & _
        "   d.ҩƷid, c.���㵥λ As �ۼ۵�λ, nvl(to_char(d.���Ч��, '9999990'),0) ���Ч��, " & _
        "   DECODE(D.ҩ�����,1,'��','��') ҩ�����, DECODE(D.ҩ������,1,'��','��') ҩ������, " & _
        "   to_char(D.����ϵ��, " & CON_FMT & ") ����ϵ��," & vbLf & _
        "   D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & vbNewLine & _
        "   D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & vbNewLine & _
        "   D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & vbNewLine & _
        "   D.ָ��������,d.ָ�����ۼ�,nvl(D.�ɱ���,0) ��ʼ�ɱ���, D.�ӳ���, q.���� ��ͬ��λ, D.ҩ�ۼ���, " & _
        "   DECODE(C.�Ƿ���,1,'��','��') ʱ��,Decode(Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '��','��') As ͣ��,d.�ϴ��ۼ�,f.���� �ϴι�Ӧ�� " & vbNewLine
    
    strSQL = strSQL & "From �շ���ĿĿ¼ C,ҩƷ��� D,�շ���Ŀ���� A,ҩƷ���� J,ҩƷ���� T,������ĿĿ¼ M,��Ӧ�� Q, ��Ӧ�� F" & vbNewLine
    
    If bln���� = False Then
        strSQL = strSQL & IIf(lng��Դ�ⷿ <> 0, " ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K", "") & vbNewLine & _
        IIf(lngĿ��ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[3] Group By ִ�п���ID,�շ�ϸĿID) I ", "") & vbNewLine
    End If
    
    strSQL = strSQL & _
        "   Where c.Id = d.ҩƷid And d.ҩ��id = t.ҩ��id And t.ҩ��id = m.Id And m.��� In ('5', '6', '7') and t.�ٴ��Թ�ҩ is null And d.ҩƷid = a.�շ�ϸĿid(+) " & _
        "     And a.����(+) = 3 And t.ҩƷ���� = j.����(+) And d.��ͬ��λid = q.Id(+) And d.�ϴι�Ӧ��id = f.id(+) "
    If bln���� = False Then
        strSQL = strSQL & IIf(lng��Դ�ⷿ <> 0, "     And D.ҩƷID=K.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        IIf(lngĿ��ⷿ <> 0, "     And D.ҩƷID=I.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "")
    End If
    
    If bln���Է������ = False Then
        strSQL = strSQL & " And" & _
             " (Decode(c.�������, 1, 1, 3, 1, 0) = " & _
             " (Select Distinct '1' From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2] And ������� In (1, 3)) Or " & _
             " Decode(c.�������, 2, 1, 3, 1, 0) =" & _
             " (Select Distinct '1' From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2] And ������� In (2, 3)) Or Exists" & _
             " (Select 1 From ��������˵�� Where �������� Like '%ҩ��' And ����id = [2])) "
    End If
    
    strSQL = strSQL & _
        IIf(bln����ͣ��ҩƷ = False, " And (C.����ʱ�� Is Null Or To_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "  (Select �շ�ϸĿid, Trim(To_Char(�ּ�, '999999999990." & String(7, "0") & "')) �ۼ� " & _
        "   From �շѼ�Ŀ Where (Sysdate Between ִ������ And ��ֹ���� or Sysdate>=ִ������ And ��ֹ���� Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine

    If byt���÷�ʽ = 1 Then
       '��������ҩ
       strSQL = strSQL & _
           "(Select a.ҩƷid,Max(�ϴβ���) AS ����, max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�, " & _
           " To_Char(Sum(b.ʵ������), '99999999999990.99') �������� " & vbNewLine & _
           "From ҩƷ��� A, ҩƷ���� B " & vbNewLine & _
           "Where a.����=1 and a.ҩƷid=b.ҩƷid And a.�ⷿid=b.�ⷿid and b.����id=[3] and b.�ڼ�=to_date(sysdate,'yyyy') "
    Else
       '��ҩ����ҩ
       strSQL = strSQL & _
           "(Select a.ҩƷid, Max(a.�ϴβ���) AS ����,max(a.ԭ����) as ԭ����, Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�, " & _
           " '' �������� " & vbNewLine & _
           "From ҩƷ��� A " & vbNewLine & _
           "Where ����=1 "
    End If
    If lng��Դ�ⷿ <> 0 Or lngĿ��ⷿ <> 0 Then
       strSQL = strSQL & " And a.�ⷿID=" & IIf(lng��Դ�ⷿ = 0, "[3]", "[2]")
    End If
    strSQL = strSQL & vbNewLine & _
       "Group By a.ҩƷid) S," & vbNewLine & _
       "(Select ҩƷID,�ⷿID,�ⷿ��λ,���ñ�־ From ҩƷ�����޶� Where �ⷿID=" & IIf(byt�༭ģʽ = 2, "[2]", "[3]") & ") E, �շ���Ŀ���� F " & vbNewLine & _
       "Where D.ҩƷID=P.�շ�ϸĿID And D.ҩƷID=S.ҩƷID" & IIf(Not (IntStockCheck = 2 And byt�༭ģʽ = 2) Or byt�̵㵥�� = 1 Or Not bln�����, "(+)", "") & _
       "  And D.ҩƷID=E.ҩƷID(+) And d.ҩƷid = f.�շ�ϸĿid(+) " & vbNewLine & _
       "Group By d.����,d.ҩ������, d.ͨ������, d.ҩƷ��Դ , d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, f.����, d.��Ʒ��, d.���, d.����" & vbNewLine & _
       ", Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����) , d.ҩ��id, d.ҩƷid,trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "'))" & vbNewLine & _
       ", trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, Nvl(d.�ϴ��ۼ�,p.�ۼ�), s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPriceDigit, "0") & "'))" & vbNewLine & _
       ", d.�ۼ۵�λ, d.����ϵ��, d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ,d.ҩ���װ,trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "'))" & vbNewLine & _
       ", s.�������,s.�����, s.�����, d.���Ч�� , d.ҩ�����, d.ҩ������, d.ʱ��,trim(to_char(d.ָ��������* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) " & vbNewLine & _
       ", trim(to_char(d.ָ�����ۼ�* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')),d.�ӳ���, e.�ⷿ��λ, d.��׼�ĺ�, s.�������" & vbNewLine & _
       ", s.��������, d.��ͬ��λ, d.ҩ�ۼ���,e.���ñ�־,d.ͣ��,d.�ϴι�Ӧ�� " & vbNewLine & _
       "Order By D.ҩ������,D.ҩƷ���� "
    Set grsMasterInput = zlDatabase.OpenSQLRecord(strSQL, "ҩƷ���", gstrNodeNo, lng��Դ�ⷿ, lngĿ��ⷿ, IIf(gint���뷽ʽ = 0, 1, 2))
    
    '*ҩƷ����*'
    If byt�༭ģʽ = 2 Then
        str�̵�sql = "Select 2 Rid,p.���� �ⷿ, k.ҩƷid, k.����, To_Char(b.�������, 'YYYY-MM-DD') As �������, k.����, k.��������, k.����,Decode(k.ԭ����, Null, d.ԭ����, k.ԭ����) as ԭ����, k.�ɱ���, k.�ۼ�, k.ʱ��, d.���ﵥλ," & vbNewLine & _
                    "       To_Char(d.�����װ, '999999999990.99999') �����װ, d.סԺ��λ, To_Char(d.סԺ��װ, '999999999990.99999') סԺ��װ, d.ҩ�ⵥλ," & vbNewLine & _
                    "       To_Char(d.ҩ���װ, '999999999990.99999') ҩ���װ,k.��Ч��, k.ʵ������, k.��������, k.�������," & vbNewLine & _
                    "                       k.�����, k.�����, k.�ϴι�Ӧ��id, k.��׼�ĺ�,f.���� ��Ӧ��" & vbNewLine & _
                    "From (Select a.�ⷿid, a.ҩƷid, nvl(a.����,0) ����, Max(a.����) ����, Max(To_Char(a.��������, 'YYYY-MM-DD')) ��������, Max(a.����) ����,Max(a.ԭ����) ԭ����, min(a.�ɱ���) �ɱ���, Avg(a.���ۼ�) �ۼ�," & vbNewLine & _
                    "              Avg(Nvl(a.���ۼ�, a.���۽�� / Decode(Nvl(a.ʵ������, 0), 0, 1, a.ʵ������))) ʱ��, Min(a.Ч��) ��Ч��," & vbNewLine & _
                    "              Sum(-1 * a.���ϵ�� * a.���� * a.ʵ������) ʵ������, Sum(-1 * a.���ϵ�� * a.���� * a.ʵ������) ��������, Sum(-1 * a.���ϵ�� * a.���� * a.ʵ������) �������," & vbNewLine & _
                    "              Sum(a.���� * a.���۽��) �����, Sum(a.���� * a.���) �����, Max(a.��ҩ��λid) �ϴι�Ӧ��id, Max(a.��׼�ĺ�) ��׼�ĺ�" & vbNewLine & _
                    "       From ҩƷ�շ���¼ A" & vbNewLine & _
                    "       Where a.���� In (1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12) And" & vbNewLine & _
                    "             a.������� > [2] " & vbNewLine & _
                    "       Group By a.�ⷿid, a.ҩƷid, nvl(a.����,0)) K, ���ű� P, ҩƷ��� D, ҩƷ�����Ϣ B, ��Ӧ�� F" & vbNewLine & _
                    "Where k.�ⷿid = p.Id And d.ҩƷid = k.ҩƷid And k.ҩƷid = b.ҩƷid(+) And k.�ⷿid = b.�ⷿid(+) And" & vbNewLine & _
                    "      k.�ϴι�Ӧ��id = f.id(+) and  k.���� = nvl(b.����(+),0) and k.������� <> 0 And k.�ⷿid = [1] "
    
        strSQL = _
            "Select max(Rid) Rid,�ⷿ,ҩƷID,����,max(�������) �������,max(����) ����,max(��������) ��������,max(����) as ������,max(ԭ����) ԭ����,max(�ɱ���) �ɱ���,max(�ۼ�) �ۼ�,max(ʱ��) ʱ��,max(���ﵥλ) ���ﵥλ,max(�����װ) �����װ,max(סԺ��λ) סԺ��λ,max(סԺ��װ) סԺ��װ,max(ҩ�ⵥλ) ҩ�ⵥλ,max(ҩ���װ) ҩ���װ," & _
            "  max(��Ч��) ��Ч��,nvl(sum(ʵ������),0) ʵ������,nvl(sum(��������),0) ��������,nvl(sum(�������),0) �������,nvl(sum(�����),0) �����,nvl(sum(�����),0) �����,max(�ϴι�Ӧ��ID) �ϴι�Ӧ��ID,max(��׼�ĺ�) ��׼�ĺ�,Max(��Ӧ��) ��Ӧ�� " & vbLf & _
            "From (Select Distinct 2 Rid, p.���� �ⷿ, k.ҩƷid, nvl(k.����,0) ����, To_Char(b.�������, 'YYYY-MM-DD') As �������, k.�ϴ����� ����," & _
            "  To_Char(k.�ϴ���������, 'YYYY-MM-DD') ��������, k.�ϴβ��� ����, Decode(k.ԭ����, Null, d.ԭ����, k.ԭ����) as ԭ����,k.ƽ���ɱ��� as �ɱ���, " & _
            "  Decode(Nvl(k.����, 0), 0, Decode(Sign(k.ʵ������), 1, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������), A.�ּ�) " & _
            "        ,Nvl(k.���ۼ�, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������) ) ) �ۼ�," & _
            "  Nvl(k.���ۼ�, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������) ) ʱ��," & _
            "  D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & _
            "  D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & _
            "  D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & _
            "  k.Ч��" & IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "-1", "") & " ��Ч��," & _
            "  k.ʵ������, k.��������, k.ʵ������ �������, k.ʵ�ʽ�� �����, k.ʵ�ʲ�� �����, k.�ϴι�Ӧ��id, k.��׼�ĺ�,f.���� ��Ӧ�� " & vbNewLine & _
            "From ���ű� P, ҩƷ��� D, ҩƷ��� K, ҩƷ�����Ϣ B, �շѼ�Ŀ A,��Ӧ�� F " & vbNewLine & _
            "Where k.�ⷿid = p.Id And d.ҩƷid = k.ҩƷid And d.ҩƷid=a.�շ�ϸĿid " & GetPriceClassString("A") & _
            "  And k.���� = 1 And k.ҩƷid = b.ҩƷid(+) And k.�ⷿid = b.�ⷿid(+) And nvl(k.����,0) = nvl(b.����(+),0) And k.�ⷿid = [1] and k.�ϴι�Ӧ��id = f.id(+) "
        If byt�̵㵥�� = 1 Then
            strSQL = strSQL & " And (K.ʵ������<>0 Or K.ʵ�ʽ��<>0 Or K.ʵ�ʲ��<>0) " & IIf(str�̵�ʱ�� <> "", vbNewLine & " union all " & vbNewLine & str�̵�sql, "") & " ) " & vbNewLine
'        ElseIf byt�̵㵥�� = 2 Then
'            '1303 ����ǿ���۵���ģ�飬��������˿������Ϊ0��ҩƷ��¼
'            gstrSQL = strSQL & " ) " & vbNewLine
        Else
            strSQL = strSQL & " And K.ʵ������<>0 " & IIf(str�̵�ʱ�� <> "", vbNewLine & " union all " & vbNewLine & str�̵�sql, "") & " ) " & vbNewLine
        End If
        If gtype_UserSysParms.P150_ҩƷ���������㷨 = 0 Then
            strSQL = strSQL & " Group By �ⷿ, ҩƷid,����" & vbNewLine & _
                    " Order By ҩƷid, ���� "
        Else
            strSQL = strSQL & " Group By �ⷿ, ҩƷid, ���� " & vbNewLine & _
                    " Order By ҩƷid, ��Ч��, ���� "
        End If

        If str�̵�ʱ�� = "" Then
            Set grsSlave = zlDatabase.OpenSQLRecord(strSQL, "ҩƷ����", IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ))
        Else
            Set grsSlave = zlDatabase.OpenSQLRecord(strSQL, "ҩƷ����", IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ), CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReleaseSelectorRS()
    If Not grsMaster Is Nothing Then
        If grsMaster.State = adStateOpen Then grsMaster.Close
        Set grsMaster = Nothing
    End If
    
    If Not grsMasterInput Is Nothing Then
        If grsMasterInput.State = adStateOpen Then grsMasterInput.Close
        Set grsMasterInput = Nothing
    End If
    
    If Not grsSlave Is Nothing Then
        If grsSlave.State = adStateOpen Then grsSlave.Close
        Set grsSlave = Nothing
    End If
End Sub


Public Sub GetPriceClass()
    '���ݵ�¼վ���ȡҩƷ�ļ۸�ȼ�
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSQL = " Select a.�۸�ȼ� " & _
            " From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B " & _
            " Where a.�۸�ȼ� = b.���� And a.���� = 0 And b.�Ƿ�����ҩƷ = 1 And a.վ�� = [1] And Nvl(b.����ʱ��, Sysdate + 1) > Sysdate "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!�۸�ȼ�
    End If
End Sub


Public Function GetPriceClassString(strTableName As String) As String
    '���ݴ����ı������ؼ۸�ȼ���������
    GetPriceClassString = " And " & IIf(strTableName = "", "�۸�ȼ� Is Null ", strTableName & ".�۸�ȼ� Is Null ")
    
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    '3 blnMoveSpace���Ƿ�ȥ���ַ��еĿո�Ture-ȥ���ո�ע��ͷβ�ո�Ĭ��ȥ��
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intAsc = Asc(Mid(strText, n, 1))
        Select Case intAsc
            Case 8, 9, 10, 13
            Case 32
                '�ո���
                If blnMoveSpace = False Then
                    strTmp = strTmp & Mid(strText, n, 1)
                End If
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function
