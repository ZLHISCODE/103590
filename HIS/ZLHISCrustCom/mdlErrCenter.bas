Attribute VB_Name = "mdlErrCenter"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/1/14
'ģ��           mdlErrCenter
'˵��           �ó����Ǻ�̨���򣬲��ᵯ��������ʾ�������Ҫ�������¼��
'==================================================================================================
'Private mobjLog                 As clsLog
Private Const M_MAX_LOG_COUNT   As Long = 8000                      '��־����¼8000�У�ÿ�ο�ʼ�����������8000�����
Private mlngRecCount            As Long                             '��־��¼����
Private Const mlngStackLen      As Integer = 40                     '���ö�ջ�ĳ���
Private mcllMethodStack         As New Collection                   '���ö�ջ����
Private mstrText                As String
Private mlngIndex               As Integer
'--------------------------------------------------------------------------------------------------
'����           LogName
'����           ������־����
'����ֵ
'����б�:
'������         ����                    ˵��
'strLogName     String                  ���ε���������־�������������Ϊ�գ��ر���־
'-------------------------------------------------------------------------------------------------
Public Property Let LogName(strLogName As String)
'    If mobjLog Is Nothing Then
'        If strLogName <> "" Then
'            Set mobjLog = New clsLog
'            Call mobjLog.OpenLog(strLogName, , False)
'        End If
'    ElseIf strLogName = "" Then
'        mobjLog.CloseLog
'        Set mobjLog = Nothing
'    End If
End Property
'--------------------------------------------------------------------------------------------------
'����           ErrorCenter
'����           ����������
'����ֵ         Integer                 0-���Լ���ִ�У�1-����(Resume),2-��ֹ����
'����б�:
'������         ����                    ˵��
'strMethod      String                  �������Ĺ���
'-------------------------------------------------------------------------------------------------
Public Function ErrorCenter(Optional ByRef strMethod As String) As Integer
'    mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strMethod, Err.Number & "-" & Err.Description
'    mobjLog.WriteListTitle String((mcllMethodStack.Count - 1) * 2, " ") & "��" & "���ö�ջ��"
'    For mlngIndex = 1 To mcllMethodStack.Count
'        mobjLog.WriteList String((mcllMethodStack.Count - 1) * 2, " ") & "��" & mcllMethodStack(mlngIndex)
'    Next
'    Call PopMethod(strMethod)
'    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'����           WarnInfo
'����           ���洦�����ִ������ֻ���ɾ��棬�������д��󲶻�
'����ֵ
'����б�:
'������         ����                    ˵��
'strWarnInfo    String                  ������Ϣ
'strMethod      String                  �������Ĺ���
'-------------------------------------------------------------------------------------------------
Public Sub LogInfo(ByRef strWarnInfo As String, ParamArray arrPars() As Variant)
'    Dim arrInfo()       As Variant
'    arrInfo = arrPars
'    mobjLog.WriteOperateArray String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strWarnInfo, arrInfo()
End Sub
'--------------------------------------------------------------------------------------------------
'����           PushMethod
'����           �����÷��������ջ
'����ֵ
'����б�:
'������         ����                    ˵��
'strMethod      String                  ������
'arrPars        String                  �����б�
'-------------------------------------------------------------------------------------------------
Public Sub PushMethod(ByRef strMethod As String, ParamArray arrPars() As Variant)
'    mstrText = ""
'    For mlngIndex = LBound(arrPars) To UBound(arrPars)
'        mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
'    Next
'    mstrText = Mid(mstrText, 2)
'    With mcllMethodStack
'        If .Count = 0 Then
'            If mstrText = "" Then
'                .Add strMethod
'            Else
'                .Add strMethod & "(" & mstrText & ")"
'            End If
'        Else
'            If mstrText = "" Then
'                .Add strMethod, , 1
'            Else
'                .Add strMethod & "(" & mstrText & ")", , 1
'            End If
'        End If
'        If .Count > mlngStackLen Then .Remove .Count
'    End With
'    If mstrText = "" Then
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strMethod
'    Else
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strMethod & "(" & mstrText & ")"
'    End If
End Sub
'--------------------------------------------------------------------------------------------------
'����           PopMethod
'����           ���������ջ�ķ����Ƴ������߽�ָ������֮ǰ���ջ�ķ����Ƴ�������ָ���ķ�����
'����ֵ
'����б�:
'������         ����                    ˵��
'strMethod      String                  �������ƣ�����ʱ����������ջ�ķ���
'-------------------------------------------------------------------------------------------------
Public Sub PopMethod(ByRef strMethod As String, ParamArray arrPars() As Variant)
'    mstrText = ""
'    For mlngIndex = LBound(arrPars) To UBound(arrPars)
'        mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
'    Next
'
'    If mstrText = "" Then
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strMethod
'    Else
'        mstrText = Mid(mstrText, 2)
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "��" & strMethod & "(" & mstrText & ")"
'    End If
'    With mcllMethodStack
'        If strMethod <> "" Then
'            For mlngIndex = 1 To .Count
'                If mcllMethodStack(mlngIndex) Like strMethod & "*" Then
'                    Exit For
'                End If
'            Next
'            If mlngIndex > .Count Then
'                If .Count > 0 Then  'û���ҵ��κ�ƥ�䣬��ɾ��һ������
'                    mlngIndex = 1
'                Else                'û��������ɾ��
'                    mlngIndex = 0
'                End If
'            End If
'        Else
'            mlngIndex = 1  '������ֻɾ��һ��
'        End If
'
'        Do While mlngIndex > 0
'            .Remove 1
'            mlngIndex = mlngIndex - 1
'        Loop
'        mlngIndex = 1
'    End With
End Sub
