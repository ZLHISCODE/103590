VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cPatients"
Option Explicit

'��������ֵ�ľֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar���� As String '�ֲ�����

Private mvar���˿��� As String '�ֲ�����
Private mvar����� As String '�ֲ�����
Private mvar��λ�� As String '�ֲ�����
Private mvar����ʱ�� As Date '�ֲ�����
Private mvar��λ�� As String '�ֲ�����
Private mvar˳��� As String '�ֲ�����
Private mvar�Ŷ�״̬ As String '�ֲ�����
Private mvar���￨�� As String '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvar������� As String '�ֲ�����
Private mvarҽ�� As String '�ֲ�����
Private mvar�Һŵ� As String '�ֲ�����
Private mvar��Ȩ�� As Long
Private mvar���б�־ As Integer
Private mvar�Һ�ʱ�� As Date
Private mvar����֤��  As String
Private mvarIC���� As String

Private mvar��ʼʱ�� As Date
Private mvar��ʼ����Ա As String
Private mvar����ʱ�� As Date
Private mvar��������Ա As String
Private mvar����̨ As String
Private mlng����ID As Long
Private mbyt������Դ As Byte

Private mvarKey As String

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let �Һŵ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Һŵ� = 5
    mvar�Һŵ� = vData
End Property

Public Property Get �Һŵ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�Һŵ�
    �Һŵ� = mvar�Һŵ�
End Property

Public Property Let ����ID(ByVal lngVar As Long)
'����Ϊ�Һ�ID����������Ϊ��ҳID
    mlng����ID = lngVar
End Property
Public Property Get ����ID() As Long
    ����ID = mlng����ID
End Property

Public Property Let ������Դ(ByVal bytVar As Byte)
'0-���1-��������
    mbyt������Դ = bytVar
End Property
Public Property Get ������Դ() As Byte
    ������Դ = mbyt������Դ
End Property

Public Property Get ��Ȩ��() As Long
    ��Ȩ�� = mvar��Ȩ��
End Property

Public Property Let ��Ȩ��(ByVal vData As Long)
    mvar��Ȩ�� = vData
End Property

Public Property Let ҽ��(ByVal vData As String)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ�� = 5
    mvarҽ�� = vData
End Property


Public Property Get ҽ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҽ��
    ҽ�� = mvarҽ��
End Property



Public Property Let �������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property


Public Property Get �������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�������
    ������� = mvar�������
End Property



Public Property Let �ѱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ѱ� = 5
    mvar�ѱ� = vData
End Property


Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�ѱ�
    �ѱ� = mvar�ѱ�
End Property



Public Property Let ���￨��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���￨�� = 5
    mvar���￨�� = vData
End Property


Public Property Get ���￨��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���￨��
    ���￨�� = mvar���￨��
End Property



Public Property Let �Ŷ�״̬(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��״̬ = 5
    mvar�Ŷ�״̬ = vData
End Property


Public Property Get �Ŷ�״̬() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��״̬
    �Ŷ�״̬ = mvar�Ŷ�״̬
End Property

Public Property Let ˳���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.˳��� = 5
    mvar˳��� = vData
End Property


Public Property Get ˳���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.˳���
    ˳��� = mvar˳���
End Property



Public Property Let ��λ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��λ�� = 5
    mvar��λ�� = vData
End Property


Public Property Get ��λ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��λ��
    ��λ�� = mvar��λ��
End Property

Public Property Let ����ʱ��(ByVal vData As Date)
    mvar����ʱ�� = vData
End Property


Public Property Get ����ʱ��() As Date
    ����ʱ�� = mvar����ʱ��
End Property

Public Property Let �Һ�ʱ��(ByVal vData As Date)
    mvar�Һ�ʱ�� = vData
End Property


Public Property Get �Һ�ʱ��() As Date
    �Һ�ʱ�� = mvar�Һ�ʱ��
End Property


Public Property Let �����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvar����� = vData
End Property


Public Property Get �����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�����
    ����� = mvar�����
End Property


Public Property Let ���˿���(ByVal vData As String)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���˿��� = 5
    mvar���˿��� = vData
    Exit Property
End Property


Public Property Get ���˿���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���˿��� = mvar���˿���
    Exit Property
End Property

Public Property Let ����(ByVal vData As String)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
    Exit Property
End Property


Public Property Get ����() As String

'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
    Exit Property
End Property



Public Property Let �Ա�(ByVal vData As String)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
    Exit Property
End Property


Public Property Get �Ա�() As String

'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�Ա�
    �Ա� = mvar�Ա�
    Exit Property
End Property



Public Property Let ����(ByVal vData As String)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
    Exit Property
End Property


Public Property Get ����() As String

'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
    Exit Property
End Property



Public Property Let ����ID(ByVal vData As Long)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
    Exit Property
End Property


Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ID
    ����ID = mvar����ID
    Exit Property
End Property

Public Function Get˳���() As String
    'ȡ��������+1
    Dim str���� As String, lng��� As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "Select To_Char(Sysdate, 'yymmdd') As ����, Max(˳���) ˳���" & vbNewLine & _
            "From �ŶӼ�¼" & vbNewLine & _
            "Where ���� Between To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And" & vbNewLine & _
            "           To_Date(To_Char(Sysdate + 1, 'yyyy-mm-dd'), 'yyyy-mm-dd')"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "cPatients.Get˳���")
    If rsTmp.RecordCount > 0 Then
        str���� = rsTmp.Fields("����")
        lng��� = Val(zlCommFun.NVL(rsTmp.Fields("˳���"), 0)) + 1
    Else
        str���� = Format(zlDatabase.Currentdate, "yyMMdd")
        lng��� = 1
    End If
    Get˳��� = str���� & "_" & Format(lng���, "000")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AddQueue(ByVal lng����ID As Long)
    Dim strSQL As String, lng˳��� As Long
    On Error GoTo errHandle
    ˳��� = Get˳���
    lng˳��� = Val(Split(˳���, "_")(1))
    strSQL = "ZL_�ŶӼ�¼_AddQueue(" & mvar����ID & "," & lng����ID & ",'" & CStr(lng˳���) & "','" & Me.�Һŵ� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.AddQueue")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function UpdateState(ByVal strState As String, ByVal lng����ID As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '�޸�״̬
    '0-δ�� 1-����Һ 2-���� 3-�˺� 4-����
    
    '5-������ 6-��ִ�� 7-ִ����
    
    '�˴����ܴ� 0-δ��,��Ϊ 1-����Һ��,������AddQueue�������
    Dim intState As Integer, lng˳��� As Long, strSQL As String
    Dim strInfo As String, blnCurDay As Boolean '�Ƿ���
    Dim blnYes As Boolean
    On Error GoTo errHandle
    
    If Me.�Ŷ�״̬ = strState Then Exit Function    '״̬��ͬ�������޸�
    
    intState = Val(Mid(strState, 1, 1))
    blnCurDay = Format(Me.����ʱ��, "yyyy-MM-dd") = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Select Case intState
        
    Case 1  '1-����Һ
        blnYes = False
        'If (Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 3 Or Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 1) And blnCurDay Then
        If Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 1 And blnCurDay Then
            '1-����Һ ����ֱ�Ӹĳ� 1������Һ
            If blnShowMsg Then MsgBox "������[" & Me.�Ŷ�״̬ & "]����Ϊ[1-����Һ]״̬��", vbQuestion, "�޸��Ŷ�״̬"
            Exit Function
        Else
            '0,2,3,4,5,6,7 �ɸ�Ϊ����Һ,���ǵ���Ļ���ʲô״̬�����Ը�Ϊ����Һ
            
            If Val(Me.�Ŷ�״̬) <> 0 Then
                If blnShowMsg Then
                    If MsgBox("׼����[" & Me.���� & "]���Ŷ�״̬��[" & Me.�Ŷ�״̬ & "]����Ϊ[1-����Һ]״̬����ȷ�ϡ�", vbInformation + vbYesNo + vbDefaultButton2, "�޸��Ŷ�״̬") = vbYes Then
                        blnYes = True
                    End If
                Else
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes Then
                lng˳��� = Val(Split(Me.˳���, "_")(1))
                If Me.������Դ = 1 Then
                    '��������
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & ",1,Null," & Me.����ID & ")"
                Else
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & ",1,'" & Me.�Һŵ� & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.�Ŷ�״̬ = "1-����Һ"
                UpdateState = True
                Exit Function
            End If
        End If
    Case 2, 3, 4 '2-����
        blnYes = False
        If Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 0 Or (Val(Mid(Me.�Ŷ�״̬, 1, 1)) >= 2 And Val(Mid(Me.�Ŷ�״̬, 1, 1)) <= 4) Then
            
            If blnShowMsg Then MsgBox "������[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬��", vbQuestion, "�޸��Ŷ�״̬"
            Exit Function
        Else
            '1,5,6,7 ��������,�˺�,����
            If blnShowMsg Then
                If MsgBox("׼����[" & Me.���� & "]���Ŷ�״̬��[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬����ȷ�ϡ�", vbInformation + vbYesNo + vbDefaultButton2, "�޸��Ŷ�״̬") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes = True Then
                lng˳��� = Val(Split(Me.˳���, "_")(1))
                If Me.������Դ = 1 Then
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",Null," & Me.����ID & ")"
                Else
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",'" & Me.�Һŵ� & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                
                'strSQL = "Zl_�ŶӼ�¼_Startend(2," & lng����ID & ",'" & Me.�Һŵ� & "',sysdate,'" & UserInfo.���� & "')"
                If Me.������Դ = 1 Then
                    strSQL = "Zl_�ŶӼ�¼_Startend(2," & lng����ID & "," & Me.����ID & ",Null," & Me.����ID & ",'" & UserInfo.���� & "')"
                Else
                    strSQL = "Zl_�ŶӼ�¼_Startend(2," & lng����ID & "," & Me.����ID & ",'" & Me.�Һŵ� & "',Null,'" & UserInfo.���� & "')"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "��д����ʱ��")

                Me.�Ŷ�״̬ = strState
                UpdateState = True
                Exit Function
            End If
        End If

    Case 5, 6  '5-������(Ҫ�кŵ�)��6�����кŵ�
        'ֻ����1������ҺתΪ 5,6
        blnYes = False
        
        If Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 1 Or Val(Mid(Me.�Ŷ�״̬, 1, 1)) = 7 Then
        
            If blnShowMsg Then
                If MsgBox("׼����[" & Me.���� & "]���Ŷ�״̬��[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬����ȷ�ϡ�", vbInformation + vbYesNo + vbDefaultButton2, "�޸��Ŷ�״̬") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            If blnYes = True Then
                lng˳��� = Val(Split(Me.˳���, "_")(1))
                If Me.������Դ = 1 Then
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",Null," & Me.����ID & ")"
                Else
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",'" & Me.�Һŵ� & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.�Ŷ�״̬ = strState
                UpdateState = True
                Exit Function
            End If
        Else
            If blnShowMsg Then MsgBox "������[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬��", vbQuestion, "�޸��Ŷ�״̬"
            Exit Function
        End If
    Case 7  '7-ִ����
        blnYes = False
        
        If InStr("0,1,2,3,4,5,6,", Val(Mid(Me.�Ŷ�״̬, 1, 1))) > 0 Then
            
            If blnShowMsg Then
                If MsgBox("׼����[" & Me.���� & "]���Ŷ�״̬��[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬����ȷ�ϡ�", vbInformation + vbYesNo + vbDefaultButton2, "�޸��Ŷ�״̬") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes = True Then
                lng˳��� = Val(Split(Me.˳���, "_")(1))
                If Me.������Դ = 1 Then
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",Null," & Me.����ID & ")"
                Else
                    strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",'" & Me.�Һŵ� & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.�Ŷ�״̬ = strState
                UpdateState = True
                Exit Function
            End If
        Else
            If blnShowMsg Then MsgBox "������[" & Me.�Ŷ�״̬ & "]����Ϊ[" & strState & "]״̬��", vbQuestion, "�޸��Ŷ�״̬"
            Exit Function
        End If
        
    Case Else
        If blnShowMsg Then MsgBox "�Ƿ���״ֵ̬��", vbQuestion, "�޸��Ŷ�״̬", gstrSysName
        Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateSequence(ByVal lng����ID As Long) As Boolean
    Dim intState As Integer, lng˳��� As Long, strSQL As String
    On Error GoTo errHandle
    
    intState = Val(Mid(Me.�Ŷ�״̬, 1, 1))
    If intState = 1 Then
        lng˳��� = Val(Split(Me.˳���, "_")(1))
        If Me.������Դ = 1 Then
            strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",Null," & Me.����ID & ")"
        Else
            strSQL = "ZL_�ŶӼ�¼_Update(" & Me.����ID & "," & lng����ID & "," & lng˳��� & "," & Me.��Ȩ�� & "," & intState & ",'" & Me.�Һŵ� & "', Null)"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSequence")
        UpdateSequence = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Public Property Let ���б�־(ByVal vData As Integer)

'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar���б�־ = vData
    Exit Property
End Property

Public Property Get ���б�־() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ID
    ���б�־ = mvar���б�־
    Exit Property
End Property

Public Property Let ����֤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvar����֤�� = vData
End Property


Public Property Get ����֤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�����
    ����֤�� = mvar����֤��
End Property

Public Property Let IC����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvarIC���� = vData
End Property


Public Property Get IC����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�����
    IC���� = mvarIC����
End Property

Public Property Let ��ʼ����Ա(ByVal vData As String)
    mvar��ʼ����Ա = vData
End Property
Public Property Get ��ʼ����Ա() As String
    ��ʼ����Ա = mvar��ʼ����Ա
End Property

Public Property Let ��������Ա(ByVal vData As String)
    mvar��������Ա = vData
End Property
Public Property Get ��������Ա() As String
    ��������Ա = mvar��������Ա
End Property

Public Property Let ����̨(ByVal vData As String)
    mvar����̨ = vData
End Property
Public Property Get ����̨() As String
    ����̨ = mvar����̨
End Property

Public Property Let ��ʼʱ��(ByVal vData As Date)
    mvar��ʼʱ�� = vData
End Property
Public Property Get ��ʼʱ��() As Date
    ��ʼʱ�� = mvar��ʼʱ��
End Property

Public Property Let ����ʱ��(ByVal vData As Date)
    mvar����ʱ�� = vData
End Property
Public Property Get ����ʱ��() As Date
    ����ʱ�� = mvar����ʱ��
End Property

Public Function SetCallTag(ByVal lng����ID As Long) As Boolean
    Dim strSQL As String
    On Error GoTo errHandle
    
    If Me.������Դ = 1 Then
        strSQL = "Zl_�ŶӼ�¼_Setcalltag(" & Me.����ID & "," & lng����ID & ",Null, " & Me.����ID & ")"
    Else
        strSQL = "Zl_�ŶӼ�¼_Setcalltag(" & Me.����ID & "," & lng����ID & ",'" & Me.�Һŵ� & "', Null)"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.SetCallTag")
    
    SetCallTag = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

