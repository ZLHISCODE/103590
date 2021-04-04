Attribute VB_Name = "mdlAppTool"
Option Explicit
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const SW_SHOWNORMAL = 1

Public Type ChooseColorType
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "User32" (ByVal hWnd As Long) As Long

Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gclsAppTool As clsAppTool       '��ǰAPPTool����
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstr��λ���� As String
Public gstrSQL As String
Public gstrMenuSys As String                '��ǰ�û�ʹ�õĲ˵�ϵͳ
Public glngSys As Long                      '��ǰϵͳ

'��������ϢϵͳҪ�õ���ȫ�ֱ���
Public gfrmMain As Object                   '����̨���ڣ���Ҫ��������Ϣ�༭���ڵĸ�����
Public gblnMessageShow As Boolean           '˵����Ϣ�������Ƿ��Ѿ���ʾ
Public gblnMessageGet  As Boolean           '˵������̨�Ƿ�Ҫ��֪ͨ���ʼ�

Public Const glngLBound As Long = 99
Public Const glngUBound As Long = 240

Public Sub GetUserInfo()
'����:�õ��û�����Ϣ

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.id,P.���,P.����,P.����,D.���� as ���ű���,D.���� as ��������,M.����ID" & _
                " from �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա M " & _
                " Where U.��Աid = P.id And P.ID=M.��ԱID and  M.ȱʡ=1 and M.����id = D.id and (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) And U.�û���=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)          '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("���ű���").Value        '��ǰ�û�
            gstrDeptName = .Fields("��������").Value        '��ǰ�û�
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    Dim strIme As String
    
    varIME = OS.SystemImes
    If Not IsArray(varIME) Then
        MsgBox "�㻹û��װ�κκ������뷨������ʹ�ñ����ܡ�" & vbCrLf & _
               "���뷨�İ�װ���ڿ����������ɡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "���Զ�����"
    strIme = zlDatabase.GetPara("���뷨")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.ListIndex = i + 1
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function IsCheckConstraint(ByVal strOwner As String, ByVal strTableName As String, ByVal strColumnName As String, ByVal bytType As Byte) As Boolean
'��ȡCheckԼ������
'bytType
'  1: �Ƿ�Ϊ Check In (0,1) Լ��
'  2: �Ƿ�Ϊ Check Is Not Null Լ��
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrH
    strTmp = "Select A.Search_Condition from All_Constraints A, All_Cons_Columns B " _
           & "Where A.Constraint_Name = B.Constraint_Name and A.owner=[1] and a.Table_Name=[2] and B.Column_Name=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "", strOwner, strTableName, strColumnName)
    If Not rsTmp.EOF And IsNull(rsTmp!search_condition) = False Then
        Select Case bytType
            Case 1: If InStr(rsTmp!search_condition, "(0,1)") Or InStr(rsTmp!search_condition, "(1,0)") Then IsCheckConstraint = True
            Case 2: If InStr(UCase(rsTmp!search_condition), "IS NOT NULL") Or InStr(UCase(rsTmp!search_condition), "IS NULL") And InStr(UCase(rsTmp!search_condition), "NOT") Then IsCheckConstraint = True
        End Select
    End If
    rsTmp.Close
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function IsPathProperty(strOwner As String, strTable As String) As String
'˵���������Լ���Ƿ�ָ��·��������ʱ�
'���أ��ӱ��������;��������;��������
    Dim i As Integer
    Dim bln���� As Boolean, blnID As Boolean, bln���� As Boolean, bln��� As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    IsPathProperty = ";"
    On Error GoTo errHandle
    
    Set rsTmp = zlDatabase.OpenSQLRecord("select * from " & strOwner & "." & strTable & " where rownum=0", "")
    If rsTmp Is Nothing Then Exit Function
    
    For i = 0 To rsTmp.Fields.Count - 1
        If rsTmp.Fields(i).Name = "����" Then
            bln���� = True
        ElseIf rsTmp.Fields(i).Name = "ID" Then
            blnID = True
        ElseIf rsTmp.Fields(i).Name = "����" Then
            bln���� = True
        End If
    Next
    rsTmp.Close
    If ((blnID Or bln����) And bln����) = False Then Exit Function
    
    strTmp = "Select b.Column_Name, c.Column_Name r_column_name,c.TABLE_NAME r_table_name " _
           & "From All_Constraints A, All_Cons_Columns B, All_Cons_Columns C " _
           & "Where a.Constraint_Name = b.Constraint_Name And a.r_Constraint_Name = c.Constraint_Name And a.Constraint_Type = 'R' " _
           & "  And a.owner=[1] and a.table_name=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "��ȡ���ӱ��ֶ��������", strOwner, strTable)
    Do While rsTmp.EOF = False
        If UCase(Nvl(rsTmp!column_name)) = "��ԴID" And UCase(Nvl(rsTmp!r_table_name)) = "RESOURCEINFO" Then
            '��������ΪBH�������ų�����
            IsPathProperty = ";;RESOURCEINFO"
        Else
            IsPathProperty = Nvl(rsTmp!column_name) & ";" & Nvl(rsTmp!r_column_name) & ";" & Nvl(rsTmp!r_table_name)
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
