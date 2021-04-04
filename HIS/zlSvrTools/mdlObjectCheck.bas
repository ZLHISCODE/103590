Attribute VB_Name = "mdlObjectCheck"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const GSTR_APPNAME As String = "�������޸�"    '������

'��������
Private mrsSequenceFromFile As ADODB.Recordset
Private mrsViewFromFile As ADODB.Recordset
Private mrsPackageFromFile As ADODB.Recordset
Private mrsFildFromFile As ADODB.Recordset
Private mrsConstraintFromFile As ADODB.Recordset
Private mrsIndexFromFile As ADODB.Recordset
Private mrsProcedureFromFile As ADODB.Recordset

Private mrsSequenceFromDB As ADODB.Recordset
Private mrsViewFromDB As ADODB.Recordset
Private mrsPackageFromDB As ADODB.Recordset
Private mrsFildFromDB As ADODB.Recordset
Private mrsConstraintFromDB As ADODB.Recordset
Private mrsIndexFromDB As ADODB.Recordset
Private mrsProcedureFromDB As ADODB.Recordset

Private mrsDataFormFile As New ADODB.Recordset
Private mrsDataFormDB As New ADODB.Recordset
Private mrsProData As New ADODB.Recordset
Private mstrSysName As String
Private mlngNum As Long
Private mlngProgress As Long
Private mblnIndex As Boolean
Private mblnReport As Boolean
Private mblnzlTables As Boolean
Private mblnProcedure As Boolean
Private mblnParameter As Boolean

Public Function IniFilePathRecordset() As ADODB.Recordset
'��ʼ����������·���ļ�¼��

    Set IniFilePathRecordset = New ADODB.Recordset
    With IniFilePathRecordset
        .Fields.Append "FilePath", adVarChar, 1000, adFldIsNullable
        .Fields.Append "SystemNum", adDouble, 20, adFldIsNullable
        .Fields.Append "FileName", adVarChar, 50, adFldIsNullable
        .Fields.Append "FileType", adVarChar, 50, adFldIsNullable
        .Fields.Append "FullVer", adVarChar, 50, adFldIsNullable
        .Fields.Append "�����", adDouble, 10, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function InitDataRecordset() As ADODB.Recordset
'���ܣ���ʼ����������SQL�ļ����ݱ���ļ�¼��
    
    Set InitDataRecordset = New ADODB.Recordset
    With InitDataRecordset
        .Fields.Append "���", adVarChar, 50, adFldIsNullable
        .Fields.Append "SQL", adVarChar, 2000, adFldIsNullable
        .Fields.Append "ϵͳ���", adVarChar, 50, adFldIsNullable
        .Fields.Append "���", adVarChar, 100, adFldIsNullable
        .Fields.Append "����", adVarChar, 100, adFldIsNullable
        .Fields.Append "������", adVarChar, 100, adFldIsNullable
        .Fields.Append "������", adVarChar, 2000, adFldIsNullable
        .Fields.Append "����", adVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function InitProDataRecordset() As ADODB.Recordset
'���ܣ���ʼ����������SQL�ļ����ݱ���ļ�¼��
    
    Set InitProDataRecordset = New ADODB.Recordset
    With InitProDataRecordset
        .Fields.Append "����SQL", adVarChar, 2000, adFldIsNullable
        .Fields.Append "ϵͳ����", adVarChar, 20, adFldIsNullable
        .Fields.Append "���", adVarChar, 20, adFldIsNullable
        .Fields.Append "������", adVarChar, 1000, adFldIsNullable
        .Fields.Append "���س̶�", adVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adVarChar, 1000, adFldIsNullable
        .Fields.Append "����˵��", adVarChar, 1000, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function GainData(ByRef rsSequenceFromFile As ADODB.Recordset, ByRef rsViewFromFile As ADODB.Recordset, ByRef rsPackageFromFile As ADODB.Recordset, ByRef rsFildFromFile As ADODB.Recordset, _
                        ByRef rsConstraintFromFile As ADODB.Recordset, ByRef rsIndexFromFile As ADODB.Recordset, ByRef rsProcedureFromFile As ADODB.Recordset, ByRef rsDataFormFile As ADODB.Recordset, _
                        ByVal rsSequenceFromDB As ADODB.Recordset, ByVal rsViewFromDB As ADODB.Recordset, ByVal rsPackageFromDB As ADODB.Recordset, ByVal rsFildFromDB As ADODB.Recordset, _
                        ByVal rsConstraintFromDB As ADODB.Recordset, ByVal rsIndexFromDB As ADODB.Recordset, ByVal rsProcedureFromDB As ADODB.Recordset, ByVal rsDataFormDB As ADODB.Recordset, _
                        ByVal blnIndex As Boolean, ByVal blnReport As Boolean, ByVal blnzlTables As Boolean, ByVal blnProcedure As Boolean, ByVal blnParameter As Boolean)
'��ʼ�����������������
    Set mrsSequenceFromFile = rsSequenceFromFile
    Set mrsViewFromFile = rsViewFromFile
    Set mrsPackageFromFile = rsPackageFromFile
    Set mrsFildFromFile = rsFildFromFile
    Set mrsConstraintFromFile = rsConstraintFromFile
    Set mrsIndexFromFile = rsIndexFromFile
    Set mrsProcedureFromFile = rsProcedureFromFile
    Set mrsDataFormFile = rsDataFormFile
    
    Set mrsSequenceFromDB = rsSequenceFromDB
    Set mrsViewFromDB = rsViewFromDB
    Set mrsPackageFromDB = rsPackageFromDB
    Set mrsFildFromDB = rsFildFromDB
    Set mrsConstraintFromDB = rsConstraintFromDB
    Set mrsIndexFromDB = rsIndexFromDB
    Set mrsProcedureFromDB = rsProcedureFromDB
    Set mrsDataFormDB = rsDataFormDB
    
    mblnIndex = blnIndex
    mblnReport = blnReport
    mblnzlTables = blnzlTables
    mblnProcedure = blnProcedure
    mblnParameter = blnParameter
    
End Function

Public Sub CompareCheck(ByRef lngNum As Long, ByRef strSysName As String, ByRef rsPro As ADODB.Recordset, ByRef lngProgress As Long)
'���ܣ��Աȱ��ؽű������ݿ���бȽ�
'������rsLocalObject-���ؽű����������ݣ�rsOraObject-���ݿ��ѯ������
    
    Set mrsProData = rsPro
    mstrSysName = strSysName
    mlngNum = lngNum
    mlngProgress = lngProgress

    Call CheckSequence
    Call CheckView
    Call CheckPackage
    Call CheckTable
    Call CheckConstraint
    Call CheckIndex
    Call CheckProcedure
    Call CheckBasicData
    lngProgress = mlngProgress
End Sub

Private Sub CheckBasicData()
'���ܣ�����������
    Dim strSQL As String
    Dim strFild As String
    Dim strLevel As String
    
    'ģ�����ݣ�ϵͳ�����
    mrsDataFormFile.Filter = "���='ģ��' and ϵͳ���=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "ģ������")
        mrsDataFormDB.Filter = "���='ģ��' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ���=" & mrsDataFormFile!���
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsDataFormFile!SQL, "ģ��", "����", _
                "���ݿ��и�ģ������ȱʧ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ���ģ������", "���,���⣺" & mrsDataFormFile!��� & "," & mrsDataFormFile!����)
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '�������ݣ�ϵͳ����ţ�����
    mrsDataFormFile.Filter = "���='����' and ϵͳ���=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "��������")
        mrsDataFormDB.Filter = "���='����' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ���=" & mrsDataFormFile!��� & " and ����='" & mrsDataFormFile!���� & "'"
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsDataFormFile!SQL, "����", "����", _
                "���ݿ��иù�������ȱʧ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ�����������", "���,���ܣ�" & mrsDataFormFile!��� & "," & mrsDataFormFile!����)
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '�������ݣ�ģ�飬�����ţ�ϵͳ
    mrsDataFormFile.Filter = "���='����' and ϵͳ���=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "��������")
        If mrsDataFormFile!���� = "NULL" Then
            mrsDataFormDB.Filter = "���='����' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ������=" & mrsDataFormFile!������ & " and ����='" & mrsDataFormFile!���� & "'"
        Else
            mrsDataFormDB.Filter = "���='����' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ������='" & mrsDataFormFile!������ & "' and ����='" & mrsDataFormFile!���� & "'"
        End If
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsDataFormFile!SQL, "����", "����", _
                "���ݿ��иò�������ȱʧ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ�����������", "ģ��,������,��������" & mrsDataFormFile!���� & "," & mrsDataFormFile!������ & "," & mrsDataFormFile!������)
        Else
            If mrsDataFormFile!���� = "NULL" Then
                strLevel = "����"
            Else
                strLevel = "��΢"
            End If
            If Not (strLevel = "��΢" And mblnParameter = False) Then
                If mrsDataFormDB!���� = mrsDataFormFile!���� Then
                    If Val(mrsDataFormDB!������) = Val(mrsDataFormFile!������) Then
                        If mrsDataFormDB!������ <> mrsDataFormFile!������ Then
                            If strLevel = "����" Then
                                strSQL = "Update Zlparameters Set ������ ='" & mrsDataFormFile!������ & "' Where ϵͳ =" & mlngNum & " And ģ�� is null And ������ =" & mrsDataFormFile!������
                            Else
                                strSQL = "Update Zlparameters Set ������ ='" & mrsDataFormFile!������ & "' Where ϵͳ =" & mlngNum & " And ģ�� =" & mrsDataFormFile!���� & " And ������ =" & mrsDataFormFile!������
                            End If
                            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", strLevel, _
                                "��������ͬ,������(" & mrsDataFormDB!������ & ")��ͬ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "����������", "ģ��,������,��������" & mrsDataFormFile!���� & "," & mrsDataFormFile!������ & "," & mrsDataFormFile!������)
                        End If
                    End If
                    If mrsDataFormDB!������ = mrsDataFormFile!������ Then
                        If Val(mrsDataFormDB!������) <> Val(mrsDataFormFile!������) Then
                            If strLevel = "����" Then
                                strSQL = "Update Zlparameters Set ������ ='" & mrsDataFormFile!������ & "' Where ϵͳ =" & mlngNum & " And ģ�� is null And ������ ='" & mrsDataFormFile!������ & "'"
                            Else
                                strSQL = "Update Zlparameters Set ������ ='" & mrsDataFormFile!������ & "' Where ϵͳ =" & mlngNum & " And ģ�� =" & mrsDataFormFile!���� & " And ������ ='" & mrsDataFormFile!������ & "'"
                            End If
                            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", strLevel, _
                                "��������ͬ,������(" & mrsDataFormDB!������ & ")��ͬ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "����������", "ģ��,������,��������" & mrsDataFormFile!���� & "," & mrsDataFormFile!������ & "," & mrsDataFormFile!������)
                        End If
                    End If
                End If
            End If
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '�������ݣ���ţ�ϵͳ
    If mblnReport Then
        mrsDataFormFile.Filter = "���='����' and ϵͳ���=" & mlngNum
        If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
        Do While Not mrsDataFormFile.EOF
            Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "��������")
            mrsDataFormDB.Filter = "���='����' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ����='" & mrsDataFormFile!���� & "'"
            If mrsDataFormDB.RecordCount = 0 Then
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "����", "����", _
                    "���ݿ��иñ�������ȱʧ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "�˹�����ñ���", "��š����ƣ�" & mrsDataFormFile!���� & "," & mrsDataFormFile!����)
            End If
            DoEvents
            mrsDataFormFile.MoveNext
        Loop
    End If
    If mblnzlTables Then
        '��Ŀ¼���ݣ�������ϵͳ
        mrsDataFormFile.Filter = "���='��Ŀ¼' and ϵͳ���=" & mlngNum
        If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
        Do While Not mrsDataFormFile.EOF
            Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "zlTables����")
            mrsDataFormDB.Filter = "���='��Ŀ¼' and ϵͳ���=" & mrsDataFormFile!ϵͳ��� & " and ����='" & mrsDataFormFile!���� & "'"
            If mrsDataFormDB.RecordCount = 0 Then
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsDataFormFile!SQL, "��Ŀ¼", "��΢", _
                    "���ݿ��иñ�Ŀ¼����ȱʧ������Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ�������", "����Ϊ��" & mrsDataFormFile!����)
            End If
            DoEvents
            mrsDataFormFile.MoveNext
        Loop
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckSequence()
'�������
    Dim i As Long
    Dim strName As String
    
    mrsSequenceFromFile.Filter = "ϵͳ���=" & mlngNum
    For i = 1 To mrsSequenceFromFile.RecordCount
        strName = mrsSequenceFromFile!����
        Call frmAppCheck.ShowProgress(mstrSysName, mrsSequenceFromFile.RecordCount, i, "����", strName)
        mrsSequenceFromDB.Filter = "����='" & strName & "'"
        If mrsSequenceFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsSequenceFromFile!SQL, "����", "����", _
                                "���ݿ��и����в����ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ�����", strName)
        End If
        DoEvents
        mrsSequenceFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckView()
'�����ͼ
    Dim i As Long
    Dim strName As String
    
    mrsViewFromFile.Filter = "ϵͳ���=" & mlngNum
    For i = 1 To mrsViewFromFile.RecordCount
        strName = mrsViewFromFile!����
        Call frmAppCheck.ShowProgress(mstrSysName, mrsViewFromFile.RecordCount, i, "��ͼ", strName)
        mrsViewFromDB.Filter = "����='" & strName & "'"
        If mrsViewFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsViewFromFile!SQL, "��ͼ", "����", _
                                "���ݿ��и���ͼ�����ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ���ͼ", strName)
        End If
        DoEvents
        mrsViewFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckPackage()
'����
    Dim i As Long
    Dim strName As String
    Dim strReName As String
    
    mrsPackageFromFile.Filter = "ϵͳ���=" & mlngNum
    For i = 1 To mrsPackageFromFile.RecordCount
        strName = mrsPackageFromFile!����
        Call frmAppCheck.ShowProgress(mstrSysName, mrsPackageFromFile.RecordCount, i, "��", strName)
        mrsPackageFromDB.Filter = "����='" & strName & "'"
        If mrsPackageFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsPackageFromFile!SQL, "��", "����", _
                "���ݿ��иð������ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹð�", strName)
        Else
            If mrsPackageFromDB!Status <> "VALID" Then
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsPackageFromFile!SQL, "��", "����", _
                    "���ݿ��а�������Ч״̬������Ӱ���Ʒ��ع��ܵ�����ʹ��", "�ؽ���", strName)
            End If
        End If
        DoEvents
        mrsPackageFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckTable()
'�����Լ��ֶ�
    Dim lngProgress As Long
    Dim strTableName As String

    Dim strFild As String
    Dim strFildLength As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim strSQL As String
    
    lngProgress = 1
    mrsFildFromFile.Filter = "ϵͳ���=" & mlngNum
    While Not mrsFildFromFile.EOF
        strTableName = mrsFildFromFile!����
        
        Call frmAppCheck.ShowProgress(mstrSysName, mrsFildFromFile.RecordCount, lngProgress, "�����ֶ�", strTableName)
        mrsFildFromDB.Filter = "����='" & strTableName & "'"
        
        If mrsFildFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsFildFromFile!SQL, "��", "����", _
                    "���ݿ��б����ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹñ�", strTableName)
            
            While strTableName = mrsFildFromFile!����
                DoEvents
                mrsFildFromFile.MoveNext
                lngProgress = lngProgress + 1
            Wend
        Else
            Do While strTableName = mrsFildFromFile!����
                strFild = mrsFildFromFile!�ֶ�
                mrsFildFromDB.Filter = "����='" & strTableName & "' and �ֶ�='" & strFild & "'"
                '�жϱ��и��ֶ��Ƿ����
                If mrsFildFromDB.RecordCount = 0 Then
                    If mrsFildFromFile!�ֶγ��� <> "" Then
                        strSQL = "Alter Table " & strTableName & " Add " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ����� & "(" & mrsFildFromFile!�ֶγ��� & ")"
                    Else
                        strSQL = "Alter Table " & strTableName & " Add " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ�����
                    End If
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "�ֶ�", "����", _
                                    "���ݿ��б�ĸ��ֶβ����ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ��ֶ�", strTableName & "��" & strFild)
                Else
                    '�ж��ֶ����������ݿ��Ƿ�һ��
                    If mrsFildFromFile!�ֶ����� <> mrsFildFromDB!�ֶ����� And mrsFildFromFile!�ֶ����� <> "VARCHAR" Then
                        If IsNull(mrsFildFromFile!�ֶγ���) = False And IsNull(mrsFildFromDB!�ֶ�����) = False Then
                            strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ����� & "(" & mrsFildFromFile!�ֶγ��� & ")"
                        Else
                            strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ�����
                        End If
                        mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "�ֶ�", "����", _
                                        "���ݿ��б�ĸ��ֶ�����(" & mrsFildFromDB!�ֶ����� & ")���׼��Ʒ(" & mrsFildFromFile!�ֶ����� & ")�в�һ�£��������ݿ���û�гɹ�", "�����ݿ��ֶ����͵���Ϊ���Ʒ��׼�ű�һ��", strTableName & "��" & strFild)
                    Else
                        'ֻ���NUMBER��VARCHAR2�������͵��ֶγ���
                        If mrsFildFromFile!�ֶ����� = "NUMBER" Then
                            If IsNull(mrsFildFromFile!�ֶγ���) = False And IsNull(mrsFildFromDB!�ֶ�ʵ�ʳ���) = False Then
                                varTemp = Split(mrsFildFromFile!�ֶγ���, ",")
                                strTemp = varTemp(0)
                                strFildLength = mrsFildFromDB!�ֶ�ʵ�ʳ���
                                If Val(strTemp) > Val(strFildLength) And Val(strFildLength) <> 0 Then
                                    strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ����� & "(" & mrsFildFromFile!�ֶγ��� & ")"
                                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "�ֶ�", "����", _
                                                "���ݿ��и��ֶγ���(" & strFildLength & ")�ȱ�׼��Ʒ(" & strTemp & ")�̣����ܵ������ݵĲ�������", "�����ݿ��ֶγ��ȵ���Ϊ���Ʒ��׼�ű�һ��", strTableName & "��" & strFild)
                                ElseIf Val(strTemp) < Val(strFildLength) Then
                                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "�ֶ�", "��΢", _
                                                "���ݿ��и��ֶγ���(" & strFildLength & ")�ȱ�׼��Ʒ(" & strTemp & ")����һ�㲻��Ӱ���Ʒ����ʹ��", "�˹������ű�", strTableName & "��" & strFild)
                                End If
                            End If
                        ElseIf mrsFildFromFile!�ֶ����� = "VARCHAR2" Then
                            If IsNull(mrsFildFromFile!�ֶγ���) = False And IsNull(mrsFildFromDB!�ֶ�ʵ�ʳ���) = False Then
                                strTemp = mrsFildFromFile!�ֶγ���
                                strFildLength = mrsFildFromDB!�ֶγ���
                                If Val(strTemp) > Val(strFildLength) And Val(strFildLength) <> 0 Then
                                    strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!�ֶ� & " " & mrsFildFromFile!�ֶ����� & "(" & mrsFildFromFile!�ֶγ��� & ")"
                                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "�ֶ�", "����", _
                                                "���ݿ��и��ֶγ���(" & strFildLength & ")�ȱ�׼��Ʒ(" & strTemp & ")�̣����ܵ������ݵĲ�������", "�����ݿ��ֶγ��ȵ���Ϊ���Ʒ��׼�ű�һ��", strTableName & "��" & strFild)
                                ElseIf Val(strTemp) < Val(strFildLength) Then
                                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "�ֶ�", "��΢", _
                                                "���ݿ��и��ֶγ���(" & strFildLength & ")�ȱ�׼��Ʒ(" & strTemp & ")����һ�㲻��Ӱ���Ʒ����ʹ��", "�˹������ű�", strTableName & "��" & strFild)
                                End If
                            End If
                        End If
                    End If
                End If
                DoEvents
                lngProgress = lngProgress + 1
                mrsFildFromFile.MoveNext
                If mrsFildFromFile.EOF Then Exit Do
            Loop
        End If
    Wend
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckConstraint()
'���Լ��
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    Dim varTemp As Variant
    Dim lngOra As Long
    Dim lngLocal As Long
    
    If mblnzlTables Then
        mrsConstraintFromDB.Filter = ""
        lngOra = mrsConstraintFromDB.RecordCount
    Else
        lngOra = 0
    End If
    
    mrsConstraintFromFile.Filter = "ϵͳ���=" & mlngNum
    lngLocal = mrsConstraintFromFile.RecordCount
    For i = 1 To mrsConstraintFromFile.RecordCount
        strName = mrsConstraintFromFile!����
        Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, i, "Լ��", strName)
        mrsConstraintFromDB.Filter = "����='" & strName & "'"
        If mrsConstraintFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsConstraintFromFile!SQL, "Լ��", "����", _
                "���ݿ��и�Լ�������ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹ�Լ��", strName)
        Else
            If mrsConstraintFromDB!Status <> "ENABLED" Then
                strSQL = "Alter Table " & mrsConstraintFromFile!���� & " Enable Novalidate Constraint " & strName
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "Լ��", "����", _
                    "Լ����ǰ���ڽ�ֹ״̬������Ӱ���Ʒ��ع��ܵ�����ʹ��", "�ָ�Լ��", strName)
            End If
            If mrsConstraintFromDB!�ֶ� <> mrsConstraintFromFile!�ֶ� Then
                If strName = "�������ϸ��_FK_����ID" And mrsConstraintFromFile!�ֶ� = "����ID,�嵥ID,����ID" Then
                
                Else
                    strSQL = "Alter Table " & mrsConstraintFromFile!���� & " Drop Constraint " & strName & " Cascade Drop Index"
                    strSQL = strSQL & "{JM|SQL�ָ���}" & vbNewLine & mrsConstraintFromFile!SQL
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "Լ��", "����", _
                        "���ݿ��и�Լ����(" & mrsConstraintFromDB!�ֶ� & ")���׼��Ʒ(" & mrsConstraintFromFile!�ֶ� & ")��һ�£�����Ӱ���Ʒ��ز�ѯ����", "ɾ����Լ�������ؽ�Լ��", strName)
                End If
            End If
        End If
        DoEvents
        mrsConstraintFromFile.MoveNext
    Next
    If mblnzlTables Then
        mrsConstraintFromDB.Filter = ""
        For i = 1 To mrsConstraintFromDB.RecordCount
            strName = mrsConstraintFromDB!����
            mrsDataFormDB.Filter = "���='��Ŀ¼' and ����='" & mrsConstraintFromDB!���� & "'"
            If mrsDataFormDB.RecordCount = 1 Then
                If mrsDataFormDB!ϵͳ��� = mlngNum Then
                    Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, lngLocal + i, "Լ��", strName)
                    mrsConstraintFromFile.Filter = "����='" & strName & "'"
                    If mrsConstraintFromFile.RecordCount = 0 Then
                        strSQL = "Alter Table " & mrsConstraintFromDB!���� & " Drop Constraint " & strName & " Cascade Drop Index"
                        mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "Լ��", "����", _
                            "���ݿ��д��ڣ�����Ʒ��׼�ű�û�У�����Ӱ���Ʒ��ز�ѯ����", "ɾ����Լ��", strName)
                    End If
                End If
            End If
            DoEvents
            mrsConstraintFromDB.MoveNext
        Next
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckIndex()
'���ܣ��������
'˵����
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    Dim varFild As Variant
    Dim lngLocal As Long
    Dim lngOra As Long
    Dim blnSpace As Boolean
    Dim strSpace As String
    
    If mblnzlTables Then
        mrsIndexFromDB.Filter = ""
        lngOra = mrsIndexFromDB.RecordCount
    Else
        lngOra = 0
    End If
    
    mrsIndexFromFile.Filter = "ϵͳ���=" & mlngNum
    lngLocal = mrsIndexFromFile.RecordCount
    For i = 1 To mrsIndexFromFile.RecordCount
        strName = mrsIndexFromFile!����
        Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, i, "����", strName)
        mrsIndexFromDB.Filter = "����='" & strName & "'"
        If mrsIndexFromDB.RecordCount = 0 Then
            If Not (strName Like "*PK" Or strName Like "*UQ*") Then
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, Split(mrsIndexFromFile!SQL, "||")(0), "����", "����", _
                    "���ݿ��и����������ڣ�����Ӱ���Ʒ�����ٶ�", "��Ӹ�����", strName)
            End If
        Else
            If mrsIndexFromDB!Status <> "VALID" Then
                strSQL = "Alter Index " & strName & " rebulid nologging"
                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", "����", _
                                    "���ݿ�������������Ч״̬������Ӱ���Ʒ�����ٶ�", "�ؽ�����", strName)
            End If
            If strName Like "*PK" Or strName Like "*UQ*" Then
                If mrsIndexFromDB!UNIQUENESS <> "UNIQUE" Then
                    If InStr(Split(mrsIndexFromFile!SQL, "||")(1), "INITIALLY DEFERRED") = 0 Then
                        strSQL = "Alter Table " & mrsIndexFromFile!���� & " Drop Constraint " & strName & " Cascade Drop Index"
                        strSQL = strSQL & "{JM|SQL�ָ���}" & vbNewLine & Split(mrsIndexFromFile!SQL, "||")(1)
                        mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", "����", _
                            "������Ψһ����Ӧ����������Ψһ����������Ӱ���Ʒ����", "ɾ����Ӧ��Լ�����ؽ�Լ��", strName)
                    End If
                End If
            Else
                If mrsIndexFromDB!�ֶ� <> mrsIndexFromFile!�ֶ� Then
                    strSQL = "drop index " & strName
                    strSQL = strSQL & "{JM|SQL�ָ���}" & mrsIndexFromFile!SQL
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", "����", _
                        "���ݿ��и��������ֶ�(" & mrsIndexFromDB!�ֶ� & ")���Ʒ��׼�ű�(" & mrsIndexFromFile!�ֶ� & ")��һ�£�����Ӱ��ϵͳ�����ٶ�", "ɾ�������������ؽ�����", strName)
                End If
            End If
        End If
        DoEvents
        mrsIndexFromFile.MoveNext
    Next
    
    If mblnzlTables Then
        strName = ""
        mrsIndexFromDB.Filter = ""
        For i = 1 To mrsIndexFromDB.RecordCount
            If strName <> mrsIndexFromDB!���� Then
                strName = mrsIndexFromDB!����
                If Not (strName Like "*PK*" Or strName Like "*UQ*") Then
                    mrsDataFormDB.Filter = "���='��Ŀ¼' and ����='" & mrsIndexFromDB!���� & "'"
                    If mrsDataFormDB.RecordCount = 1 Then
                        If mrsDataFormDB!ϵͳ��� = mlngNum Then
                            Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, lngLocal + i, "����", strName)
                            mrsIndexFromFile.Filter = "����='" & strName & "'"
                            If mrsIndexFromFile.RecordCount = 0 Then
                                strSQL = "drop index " & strName
                                mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", "����", _
                                    "���ݿ��д��ڣ�����Ʒ��׼�ű�û�У�����Ӱ���Ʒ��ع��ܵ�д������", "ɾ��������", strName)
                            Else
                                If mblnIndex Then
                                    blnSpace = False
                                    strSpace = ""
                                    Do While Not mrsIndexFromFile.EOF
                                        If mrsIndexFromFile!��ռ� = mrsIndexFromDB!��ռ� And mrsIndexFromFile!��ռ� <> "" Then
                                            blnSpace = True
                                        Else
                                            strSpace = mrsIndexFromFile!��ռ�
                                        End If
                                        mrsIndexFromFile.MoveNext
                                    Loop
                                    If blnSpace = False And IsNull(mrsIndexFromDB!��ռ�) = False And strSpace <> "" And mrsIndexFromDB!���� <> "��Ѫ����" Then
                                        strSQL = "Alter Index " & strName & " rebulid tablespace " & strSpace & " nologging"
                                        mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����", "��΢", _
                                            "������ռ�(" & mrsIndexFromDB!��ռ� & ")���Ʒ��׼�ű�(" & strSpace & ")��һ�£���ά���Խ���", "�ؽ�����", strName)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            mrsIndexFromDB.MoveNext
        Next
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckProcedure()
'���ܣ�������/����
'˵���������еı�ռ�ֵ�ŵĲ���λ��
    Dim i As Long
    Dim strName As String
    Dim strTemp As String
    Dim strFild As String
    Dim varDBFild As Variant
    Dim varFileFild As Variant
    Dim strSQL As String
    
    mrsProcedureFromFile.Filter = "ϵͳ���=" & mlngNum
    For i = 1 To mrsProcedureFromFile.RecordCount
        strName = mrsProcedureFromFile!����
        strFild = mrsProcedureFromFile!�ֶ�
        Call frmAppCheck.ShowProgress(mstrSysName, mrsProcedureFromFile.RecordCount, i, "����/����", strName)
        mrsProcedureFromDB.Filter = "����='" & strName & "'"
        If mrsProcedureFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, mrsProcedureFromFile!SQL, "����/����", "����", _
                "���ݿ��иù��̻��������ڣ�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "��Ӹù���/����", strName)
        Else
            If mrsProcedureFromDB!Status = "VALID" Then
                varFileFild = Split(mrsProcedureFromFile!�ֶ�, ",")
                varDBFild = Split(Nvl(mrsProcedureFromDB!�ֶ�, ""), ",")
                If UBound(varFileFild) < UBound(varDBFild) Then
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "����/����", "��΢", _
                        "���ݿ��иù��̻�����������(" & UBound(varDBFild) + 1 & "��)�ȱ�׼��Ʒ(" & UBound(varFileFild) + 1 & "��)�࣬����Ӱ���Ʒ��ع��ܵ�����ʹ��", "�˹������ű�", strName)
                ElseIf UBound(varFileFild) > UBound(varDBFild) Then
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "����/����", "����", _
                        "���ݿ��иù��̻�����������(" & UBound(varDBFild) + 1 & "��)�ȱ�׼��Ʒ(" & UBound(varFileFild) + 1 & "��)�٣�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "�˹������ű�", strName)
                Else
                    If mrsProcedureFromFile!�ֶ� <> mrsProcedureFromDB!�ֶ� Then
                        mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, "", "����/����", "����", _
                            "���ݿ��иù��̻�������˳�������(" & mrsProcedureFromDB!�ֶ� & ")���׼��Ʒ(" & mrsProcedureFromFile!�ֶ� & ")��һ�£�����Ӱ���Ʒ��ع��ܵ�����ʹ��", "�˹������ű�", strName)
                    End If
                End If
            Else
                If mblnProcedure Then
                    strTemp = UCase(Mid(mrsProcedureFromFile!SQL, 1, 50))
                    If InStr(strTemp, "PROCEDURE") > 0 Then
                        strSQL = "Alter procedure " & strName & " Compile"
                    Else
                        strSQL = "Alter Function " & strName & " Compile"
                    End If
                    mrsProData.AddNew Array("ϵͳ����", "����SQL", "���", "���س̶�", "��������", "����˵��", "������"), Array(mstrSysName, strSQL, "����/����", "����", _
                        "���ݿ��иù��̻���������Ч״̬������Ӱ���Ʒ��ع��ܵ�����ʹ��", "���±���ù���/����", strName)
                End If
            End If
        End If
        DoEvents
        mrsProcedureFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Public Function SetSelectRecordset(ByRef strSelect As String, ByRef strFilds As String, ByVal arrFields As Variant, _
     ByRef strTableName As String) As ADODB.Recordset
'���ܣ���Insert Into�����ֶ������ֶ�ֵת���ɼ�¼������
'������
'  strSelect��Insert Into���
'  strFilds���ֶ�ֵ
'  arrFields��Insert Into���ֶ�������
'  strTableName���������
'���أ���¼������
    Dim rsSelect As New ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Integer, j As Integer
    Dim arrSelect As Variant, arrValue As Variant
    Dim strTmp As String, strSQL As String, strHead As String
    Dim bytLevel As Byte
    Dim strTeam As String, strParentLevel As String
    Dim varTemp As Variant
    Dim strModifySQL As String

    strSelect = UCase(strSelect)
    
    '���ɼ�¼��
    With rsSelect
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "����SQL", adVarChar, 2000
        For i = LBound(arrFields) To UBound(arrFields)
            If (arrFields(i) Like "*��" Or arrFields(i) = "�˵�ID") And arrFields(i) <> "���" Then
                .Fields.Append Trim(arrFields(i)), adBigInt
            Else
                .Fields.Append Trim(arrFields(i)), adVarChar, 300
            End If
        Next
        .Open
    End With
    
    '��strScriptȡ���ֶ�ֵ
    If strSelect Like "*SELECT *" Then
        'select��ʽ
        lngBegin = InStr(strSelect, "SELECT ")
        arrSelect = Split(Mid(strSelect, lngBegin), "UNION ALL")
    Else
        'values��ʽ
        If strSelect Like "*VALUES*" Then
            lngBegin = InStr(strSelect, "VALUES") + 6
            lngBegin = InStr(Mid(strSelect, lngBegin), "(") + lngBegin
        Else
            Exit Function
        End If
        lngEnd = InStr(Mid(strSelect, lngBegin), ")") - 1
        arrSelect = Array()
        ReDim Preserve arrSelect(UBound(arrSelect) + 1)
        strTeam = Mid(strSelect, lngBegin, lngEnd)
        If InStr(strTeam, "(") > 0 Then strTeam = strTeam & ")"
        arrSelect(UBound(arrSelect)) = "SELECT " & strTeam & " FROM DUAL "
    End If
    '���¼��дֵ
    strHead = ""
    On Error GoTo errHandle
    For i = LBound(arrSelect) To UBound(arrSelect)
        strSQL = Trim(ClearSpace(arrSelect(i), True))
        strTmp = strSQL
        If Trim(strTmp) <> "" Then
            If strTmp Like "SELECT *[,| ]A.[*] FROM *" Then
                '����ͷ
                lngBegin = InStr(strTmp, "SELECT ") + 7
                lngEnd = InStr(Mid(strTmp, lngBegin), ",A.*") - 1
                If lngEnd < 0 Then
                    lngEnd = InStr(Mid(strTmp, lngBegin), ", A.*") - 1
                End If
                strHead = Mid(strTmp, lngBegin, lngEnd)
            ElseIf strTmp Like "*) A;*" Then
                '����β
            Else
                '����
                If strTmp Like "* DUAL*" Then
                    lngBegin = InStr(strTmp, "SELECT ") + 7
                    lngEnd = InStr(Mid(strTmp, lngBegin), " FROM ")
                    If strHead <> "" Then
                        strSQL = strHead & "," & Mid(strSQL, lngBegin, lngEnd)
                    Else
                        strSQL = Mid(strSQL, lngBegin, lngEnd)
                    End If
                    arrValue = Split(strSQL, ",")
                    strModifySQL = "INSERT INTO " & strTableName & "(" & strFilds & ") VALUES (" & strSQL & ")"
                    
                    rsSelect.AddNew
                    rsSelect!����SQL = strModifySQL
                    For j = LBound(arrValue) + 1 To UBound(arrValue) + 1
                        If rsSelect.Fields(j).Type = adVarChar Then
                            rsSelect.Fields(j).value = Trim(Replace(Trim(arrValue(j - 1)), "'", ""))
                        Else
                            If Not Trim(LCase(arrValue(j - 1))) Like "*NULL" Then
                                rsSelect.Fields(j).value = Trim(Replace(Trim(arrValue(j - 1)), "'", ""))
                            Else
                                rsSelect.Fields(j).value = 0
                            End If
                        End If
                        rsSelect.Fields(j).value = Trim(rsSelect.Fields(j).value)
                    Next
                    rsSelect.Update
                End If
            End If
        End If
    Next
    
    Set SetSelectRecordset = rsSelect
    Exit Function
    
errHandle:

End Function

Public Function ReplaceNoteMark(ByVal strScript As String, ByVal strSymbol As String, ByVal strSymbolNew As String) As String
'���ܣ��Խű��������ڵķ����滻
'������
'  strScript��Ҫ�����SQL�ű�
'  strSymbol��ָ��ԭ�ַ�
'  strSymbolNew���滻���ַ�
'���أ�

    Const STR_SQM  As String = "'"

    Dim l As Long
    Dim blnStart As Boolean
    Dim strTmp As String
    
    If strSymbol = "" Then Exit Function
    If Len(strSymbol) > 1 Then Exit Function
    For l = 1 To Len(strScript)
        If Mid(strScript, l, 1) = STR_SQM Then
            blnStart = Not blnStart
            strTmp = strTmp & Mid(strScript, l, 1)
        Else
            If Mid(strScript, l, 1) = strSymbol And blnStart Then
                strTmp = strTmp & strSymbolNew
            Else
                strTmp = strTmp & Mid(strScript, l, 1)
            End If
        End If
    Next
    
    ReplaceNoteMark = strTmp
End Function

Public Function ClearSpace(ByVal strVal As String, Optional ByVal blnSpace As Boolean = False) As String
'���ܣ��������Ŀո�
'������
'  strVal����Ҫ������ִ�
'  blnSpace�����з�ת�ո��
'���أ����������ִ�

    Dim strResult As String
    Dim l As Long
    Dim blnStart As Boolean
    
    If strVal = "" Then Exit Function
    
    '�����������ڵĻس����з�
    For l = 1 To Len(strVal)
        If Mid(strVal, l, 1) = "'" Then
            blnStart = Not blnStart
        End If
        If blnStart Then
            If Asc(Mid(strVal, l, 1)) = Asc(vbCrLf) Or Asc(Mid(strVal, l, 1)) = Asc(vbCr) Then
                strVal = Left(strVal, l - 1) & "[[ENTER]]" & Mid(strVal, l + 1)
            End If
        End If
    Next
    
    '
    strResult = Replace(Replace(strVal, vbTab, " "), vbCrLf, IIf(blnSpace, " ", ""))
    If blnSpace Then strResult = Replace(strResult, vbCr, " ")
    
    '��ԭ�س����з�
    strResult = Replace(strResult, "[[ENTER]]", vbCr)
    
    Do While InStr(strResult, "  ") > 0
        strResult = Replace(strResult, "  ", " ")
    Loop
    
    ClearSpace = strResult
    
End Function

Public Function GetSystemList() As ADODB.Recordset
'���ܣ���ȡZlsystems�и�ϵͳ����Ϣ
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ��� ϵͳ���, ���� ϵͳ����, �汾�� ϵͳ�汾��, ������ ϵͳ������, �����, ������װ From Zlsystems where Upper(������)=[1] Order by Nvl(�����,0),���"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡϵͳ�嵥��Ϣ", gstrUserName)
    
    Set GetSystemList = rsSys
End Function

Public Function GetSystemSetupIni() As ADODB.Recordset
'���ܣ���ȡZlsystems�и�ϵͳ�İ�װ�ű��ļ�λ��
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.ϵͳ ϵͳ���, A.����, upper(A.�ļ���) �ļ��� From Zlsysfiles a Where Upper(������)=[1] and  A.���� in(1,2) Order By ϵͳ,����"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡϵͳ��װ�ű��ļ�", gstrUserName)
    
    Set GetSystemSetupIni = rsSys
End Function

Public Sub ReleaseMe()
'�رռ��������ʱ�ͷ�ģ�鴰��

    Set mrsSequenceFromFile = Nothing
    Set mrsViewFromFile = Nothing
    Set mrsPackageFromFile = Nothing
    Set mrsFildFromFile = Nothing
    Set mrsConstraintFromFile = Nothing
    Set mrsIndexFromFile = Nothing
    Set mrsProcedureFromFile = Nothing
    Set mrsDataFormFile = Nothing
    
    Set mrsSequenceFromDB = Nothing
    Set mrsViewFromDB = Nothing
    Set mrsPackageFromDB = Nothing
    Set mrsFildFromDB = Nothing
    Set mrsConstraintFromDB = Nothing
    Set mrsIndexFromDB = Nothing
    Set mrsProcedureFromDB = Nothing
    Set mrsDataFormDB = Nothing
    
    Set mrsProData = Nothing
    mstrSysName = ""
End Sub

