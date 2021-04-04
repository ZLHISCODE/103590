VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'�ӿ�˵��:
'1.�շѽӿ�
'1.1 zlDeviceSetup������ִ�ж˿۷ѵ�֧����ʽ
'1.2 zlSquareAffirm:����ִ�ж˸���

'*********************************************************************************************************************************************
'�ڲ�����������˵��
'1.GetCardSquareObject:�����������
'*********************************************************************************************************************************************
Private mobjSquare As Object

Private Function GetCardSquareObject(ByRef objCardSquare_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:objCardSquare_Out-���ض���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-06-27 14:10:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjSquare Is Nothing Then
        Set objCardSquare_Out = mobjSquare: GetCardSquareObject = True
        Exit Function
    End If
    
    err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    If err <> 0 Then
        err = 0: On Error GoTo 0
         MsgBox "����zl9CardSquare����ʧ�ܣ�����ò����Ƿ���ڻ�ע��", vbInformation + vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    If mobjSquare.zlInitComponents(Nothing, glngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         MsgBox "��ʼ��ʧ��,������ػ�����������ȷ����HIS������ϵ", vbInformation + vbOKOnly, "��ʾ"
         Set mobjSquare = Nothing
         Exit Function
    End If
    
    Set objCardSquare_Out = mobjSquare: GetCardSquareObject = True
End Function

Public Function zlSquareAffirm(ByVal lng����ID As Long, ByVal strҽ����Ϣ As String, ByVal strFeeNos As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ����ִ�ж˸���
    ' ��� : strҽ����Ϣ-��ʽ��ҽ��IDs|ִ�в���ID������ִ�в��Ŷ�Ӧ��ҽ����ȡ���۵�
    '        strFeeNos-���θ��ѵĻ��۵������ҽ��Ҳ�����˻��۵��������ߺϲ�
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/6/27 09:46
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, strҽ��IDs As String
    Dim lngִ�в���ID As Long
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    If strҽ����Ϣ = "" And strFeeNos = "" Then Exit Function
    If GetCardSquareObject(mobjSquare) = False Then Exit Function
    
    strҽ��IDs = Split(strҽ����Ϣ & "|", "|")(0)
    lngִ�в���ID = Val(Split(strҽ����Ϣ & "|", "|")(1))
    strNos = strFeeNos
    
    If strҽ����Ϣ <> "" Then
        If InStr(strҽ����Ϣ, ",") > 0 Then '�ж��ҽ��ID�����
            strSQL = "Select Distinct NO From (" & vbNewLine & _
                    "With ҽ�� As (Select /*+cardinality(B,10) */" & vbNewLine & _
                    "                 Column_Value" & vbNewLine & _
                    "                From Table(f_Num2list([2])) b)" & vbNewLine & _
                    "Select a.No" & vbNewLine & _
                    "From ������ü�¼ a, ҽ�� c" & vbNewLine & _
                    "Where a.ҽ����� = c.Column_Value  And a.��¼���� = 1 And (a.��¼״̬ = 0 Or a.��¼״̬ = 1 And a.����id Is Null) And" & vbNewLine & _
                    "      a.����id = [1] And a.ִ�в���id = [3] And Nvl(a.����״̬, 0) = 0" & vbNewLine & _
                    "Union all" & vbNewLine & _
                    "Select a.No" & vbNewLine & _
                    "From ������ü�¼ a, ����ҽ����¼ b, ҽ�� c" & vbNewLine & _
                    "Where a.ҽ����� = b.ID And b.���id = c.Column_Value And a.��¼���� = 1 And (a.��¼״̬ = 0 Or a.��¼״̬ = 1 And a.����id Is Null) And" & vbNewLine & _
                    "      a.����id = [1] And a.ִ�в���id = [3] And Nvl(a.����״̬, 0) = 0)"
        Else
            strSQL = "Select Distinct No" & vbNewLine & _
                    "From ������ü�¼" & vbNewLine & _
                    "Where ҽ����� In (Select Id From ����ҽ����¼ Where Id = [2] Union All " & vbNewLine & _
                    "                   Select Id From ����ҽ����¼ Where ���id = [2]) And ��¼���� = 1 And (��¼״̬ = 0 Or ��¼״̬ = 1 And ����id Is Null) And" & vbNewLine & _
                    "      ����id = [1] And ִ�в���id = [3] And Nvl(����״̬, 0) = 0"
        End If

        Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface", lng����ID, strҽ��IDs, lngִ�в���ID)
        Do While Not rsData.EOF
            If InStr("," & strNos & ",", "," & Nvl(rsData!NO) & ",") = 0 Then
                strNos = strNos & "," & Nvl(rsData!NO)
            End If
            rsData.MoveNext
        Loop
        If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    End If
    
    If strNos = "" Then Exit Function
    
    If mobjSquare.zlSquareAffirm(Nothing, glngModule, "", lng����ID, , , 1, strNos) = False Then
        Exit Function
    End If
    zlSquareAffirm = True
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Sub zlDeviceSetup()
    '---------------------------------------------------------------------------------------
    ' ���� : ����ִ�ж˿۷ѵ�֧����ʽ
    ' ���� : ���ϴ�
    ' ���� : 2019/6/27 09:45
    '---------------------------------------------------------------------------------------
    On Error GoTo errH
    Call gzlComLib.zlCommFun.DeviceSetup(Nothing, glngSys, glngModule)
    Exit Sub
errH:
    MsgBox err.Description
End Sub

Private Sub Class_Terminate()
    If Not mobjSquare Is Nothing Then Set mobjSquare = Nothing
End Sub