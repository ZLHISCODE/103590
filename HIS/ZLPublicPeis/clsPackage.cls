VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��������
'######################################################################################################################
Private mstrSQL As String
Private mstrTitle As String

Public Function Get_PeisPersonReportBill(ByVal lngPersonKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************

    On Error GoTo errHand

    mstrSQL = _
        "Select A.��ʼʱ��, A.����ʱ��, B.��������, A.����id, A.����id" & vbNewLine & _
        "From ���������Ա A, ��������¼ B" & vbNewLine & _
        "Where A.����id = [1] And A.���״̬ > 1 And A.����id = B.ID And A.��ʼʱ��  is not Null " & vbNewLine & _
        "Order By A.��ʼʱ�� Desc"

    Set Get_PeisPersonReportBill = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngPersonKey)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function Get_PeisPersonItem(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long) As ADODB.Recordset
    
    On Error GoTo errHand
    
    mstrSQL = _
            "Select Distinct '����' As ͼ��, a.�嵥id, b.Id, b.���� As ��Ŀ, a.�ɼ���ʽid, a.�걾��̬, a.��ע˵��, c.����˳��" & vbNewLine & _
            "From ��������� A, �����ĿĿ¼ B, �����Ŀ˳�� C" & vbNewLine & _
            "Where a.����id = [1] And a.����id = [2] And a.�����Ŀid = b.Id And b.Id = c.�����Ŀid(+)" & vbNewLine & _
            "Order By b.Id, c.����˳��"



    Set Get_PeisPersonItem = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngTaskKey, lngPersonKey)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function



Public Function Get_PeisPersonConclusion(ByVal lngStation As Long, ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, Optional ByVal lng�嵥id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case lngStation
        Case 1
            mstrSQL = _
                    "Select ID, ��������, ��¼���" & vbNewLine & _
                    "From ����������" & vbNewLine & _
                    "Where �������� = 1 And ����id = [1] And ����id = [2] And �嵥id=[3] And ��¼���� = 1" & vbNewLine & _
                    "Order By ��¼���"
        Case 2
            mstrSQL = _
                    "Select ID, ��������, ��¼���" & vbNewLine & _
                    "From ����������" & vbNewLine & _
                    "Where �������� = 2 And ����id = [1] And ����id = [2] And ��¼���� = 1" & vbNewLine & _
                    "Order By ��¼���"
    End Select
    
    Set Get_PeisPersonConclusion = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngTaskKey, lngPersonKey, lng�嵥id)
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function get_PeisPersonResult(ByVal lng����id As Long, ByVal lng��Աid As Long, ByVal lng�嵥id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = _
        "Select a.�嵥id, a.�����Ŀid As ID, b.���� As ָ��, a.���, a.���� As ��ʾ, a.�ο�" & vbNewLine & _
        "From ��������� A, ���ָ��Ŀ¼ B" & vbNewLine & _
        "Where a.����id = [1] And a.����id = [2] And a.�嵥id = [3] And a.���ָ��id = b.Id"

    
    Set get_PeisPersonResult = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng����id, lng��Աid, lng�嵥id)
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function
