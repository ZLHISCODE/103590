VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer timTransmit 
      Enabled         =   0   'False
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event AfterTransmit(ByVal strLog As String, ByVal blnTransmitting As Boolean)
Private mobjComLib As Object
Private mobjMachine As Object
Private mstrBaseData As String                              '�������ݴ��ͣ����clsDataTransmit.Transmit�Ĳ���˵��
Private mclsLog As clsLog
Private mtypParams As TYPE_PARAMS
Private mblnTransmitting As Boolean                         'True���ڴ�������
Private mstrUser As String

'�������ݴ���
Public Property Get BaseData() As String
    BaseData = mstrBaseData
End Property
Public Property Let BaseData(ByVal strValue As String)
    mstrBaseData = strValue
End Property

'Ŀǰ֧�ֶ�ʱ���͵�ҵ�����ݣ���������ѡ��
'����е�����ʱҵ�����ݴ��ͣ���ͬ���޸� timTransmit_Timer() �¼���
Public Property Get SupportData() As String
    SupportData = "�����շ�|�����˷ѣ�������|���﷢ҩ|סԺ��ҩ" '& "|����״̬�����ͣ�"
End Property

Public Function ShowMe(ByVal strUser As String, ByVal objComLib As Object, ByVal clsVar As clsLog) As Boolean
    Dim strMsg As String, strTmp As String

    mstrUser = strUser
    Set mclsLog = clsVar
    
    If mobjMachine Is Nothing Then
        On Error Resume Next
        Set mobjMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err.Number <> 0 Then
            strTmp = "��ȷ��zlDrugMachine������ע�ᣡ"
            mclsLog.Add strTmp
            mclsLog.Save
            RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    strTmp = "��ʼ��ҩƷ�Զ����豸�ӿڲ�����"
    mclsLog.Add vbNewLine & "" & Now
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    On Error GoTo hErr

    Set mobjComLib = objComLib

    If mobjMachine.Init(Val("2-������"), mobjComLib, strMsg) = False Then
        mclsLog.Add strMsg
        RaiseEvent AfterTransmit(strMsg, mblnTransmitting)
        strTmp = "��ʼ��zlDrugMachine����ʧ�ܣ�"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        Exit Function
    End If
    mclsLog.Save

    Me.Show
    Me.Visible = False
    
    ShowMe = True

    Exit Function
    
hErr:
    strTmp = Err.Description
    mclsLog.Add strTmp
    mclsLog.Save
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
End Function

Public Sub ReadParams()
'���ܣ���ȡ�����������浽����
    
    Dim objXML As New clsXML
    Dim strFile As String, strTmp As String
    
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, App.EXEName, "") & "zlDrugMachineManage\zlDrugMachine.cfg"
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        strTmp = "�����ߵĲ����ļ�����ȷ��"
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    End If

    With mtypParams
        .��ʱ���� = Val(GetParameter(objXML, "cycle", "5"))
        .��Ч���� = Val(GetParameter(objXML, "validdays", "2"))
        .��ʾ������� = Val(GetParameter(objXML, "viewlines", "200"))
        .�����־ = Val(GetParameter(objXML, "output", "0")) = 1
        .��ϸ��־ = Val(GetParameter(objXML, "detailed", "0")) = 1
        .������־���� = Val(GetParameter(objXML, "savedays", "7"))
        .ҵ������ = Trim(GetParameter(objXML, "businessdata", ""))
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjMachine = Nothing
    Set mobjComLib = Nothing
End Sub

Private Sub timTransmit_Timer()
    Dim strSQL As String, strTmp As String, strINF As String
    Dim rsTemp As ADODB.Recordset

    '����Timer���֧��65535���룬��ˣ�ͨ����ͨ��ʽʵ�ִ�С65��Ķ�ʱ����
    If (Timer() - Val(Me.Tag)) \ 60 < mtypParams.��ʱ���� Then Exit Sub
    
    timTransmit.Tag = "1"   '��ʼ��ʱ����
    timTransmit.Enabled = False
    
    On Error GoTo hErr
    
    mclsLog.Add vbNewLine & "" & Now
    
    mblnTransmitting = True
    
    strTmp = "��ʼҵ�����ݶ�ʱ����"
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    strTmp = "|" & mtypParams.ҵ������ & "|"
    
    strSQL = _
        "Select Distinct a.NO ������, a.����, a.�ⷿid  " & vbNewLine & _
        "From ҩƷ�շ���¼ A, ҩƷ�շ������־ B " & vbNewLine & _
        "Where a.No = b.������ And a.���� = b.���� " & vbNewLine & _
        "    And a.�������� Between Sysdate - [2] And Sysdate And b.ҵ����� = [1] And Instr(';0;11;12;', ';' || Nvl(b.��־, 0) || ';') > 0 "
    
    '�����շ�
    If InStr(strTmp, "|�����շ�|") > 0 Then
        Call TransBusiness(21, strSQL, "1-�����շ�")
    End If
    
    '�����˷ѣ�������
    If InStr(strTmp, "|�����˷ѣ�������|") > 0 Then
        Call TransBusiness(25, strSQL, "2-�����˷ѣ�������")
    End If
    
    '���﷢ҩ
    If InStr(strTmp, "|���﷢ҩ|") > 0 Then
        Call TransBusiness(22, strSQL, "3-���﷢ҩ")
    End If
    
    'סԺ��ҩ
    If InStr(strTmp, "|סԺ��ҩ|") > 0 Then
        strSQL = _
            "Select b.�շ�id " & vbNewLine & _
            "From ҩƷ�շ���¼ A, ҩƷ�շ�סԺ��־ B " & vbNewLine & _
            "Where a.Id = b.�շ�id And a.�������� Between Sysdate - [2] And Sysdate And b.ҵ����� = [1] And b.��־ > 10 "

        Call TransBusiness(21, strSQL, "4-סԺ��ҩ")
    End If
    
    '�����豸��Ҫ��ҩ������״̬֪ͨ
    If InStr(strTmp, "|����״̬�����ͣ�|") > 0 Then
        '����ҩ���ķ�ҩ����״̬
        strSQL = _
            "Select f_List2str(Cast(Collect(���) As t_Strlist), ';') �ӿڱ�� " & vbNewLine & _
            "From ҩƷ�豸�ӿ� " & vbNewLine & _
            "Where ͣ������ Is Null "
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡҩƷ�豸�ӿ�")
        If rsTemp.RecordCount > 0 Then
            strINF = IIf(IsNull(rsTemp!�ӿڱ��), "", rsTemp!�ӿڱ��)
        End If
        rsTemp.Close

        Call TransBase(strINF & "|5")
    End If
    
    '������޻������ݴ���
    If mstrBaseData <> "" Then
        strTmp = "��ʼ���ͻ�������"
        mclsLog.Add strTmp
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
        Call TransBase(mstrBaseData)
        
        mblnTransmitting = False
        strTmp = "��ɴ��ͻ�������"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
        mstrBaseData = ""
    Else
        mblnTransmitting = False
        strTmp = "���ҵ�����ݶ�ʱ����"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    End If
    
    Me.Tag = Timer()
    timTransmit.Enabled = True
    timTransmit.Tag = ""    '��ɶ�ʱ����
    
    Exit Sub
    
hErr:
    mblnTransmitting = False
    timTransmit.Enabled = True
    timTransmit.Tag = ""
End Sub

Private Sub TransBusiness(ByVal intType As Integer, ByVal strSQL As String, ByVal strInfo As String)
'���ܣ�����ҵ������
'������
'  intType��ҵ�����
'           21-��ҩ[�����סԺ������ϸ�ϴ�]��
'           22-��ʼ��ҩ��
'           23-��ɷ�ҩ��
'           25-����������ҩ��
'  strInfo��1-�����շѣ�2-�����˷ѣ���������3-���﷢ҩ��4-סԺ��ҩ
    
    Dim rsTemp As ADODB.Recordset
    Dim strData As String, strBill As String, strMsg As String, strTmp As String
    
    On Error GoTo hErr
    
    strTmp = "��ȡ��" & strInfo & "������"
    mclsLog.Add strTmp
    mclsLog.Add strSQL, 1, Val("1-��ϸ��־")
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
'    With mclsLog
'        .Add "��ѯ������", 1, Val("1-��ϸ��־")
'        .Add strInfo, 1, Val("1-��ϸ��־")
'        .Add mtypParams.��Ч����, 1, Val("1-��ϸ��־")
'    End With
    
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, strInfo, intType - 20, mtypParams.��Ч����)
    Do While rsTemp.EOF = False
        If Val(strInfo) = Val("2-�����˷ѣ�������") Then
            strBill = strBill & ";" & rsTemp!���� & "," & rsTemp!������ & "," & IIf(IsNull(rsTemp!�ⷿid), "", rsTemp!�ⷿid)
        Else
            strBill = strBill & ";" & rsTemp!���� & "," & rsTemp!������
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    mclsLog.Add strBill, 1, Val("1-��ϸ��־")
    
    If Left(strBill, 1) = ";" Then strBill = Mid(strBill, 2)
    
    Select Case Val(strInfo)
    Case Val("1-�����շ�"), Val("3-���﷢ҩ")
        strData = "1|" & strBill
        
    Case Val("2-�����˷ѣ�������")
        strData = strBill
        
    Case Val("4-סԺ��ҩ")
        strData = "2|" & strBill
        
    End Select
    
    If strBill = "" Then
        strTmp = "��" & strInfo & "����������"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        Exit Sub
    End If
    
    strTmp = "��ʼ���͡�" & strInfo & "������"
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    '��������
    If mobjMachine.Operation(mstrUser, intType, strData, strMsg) Then
        '������־���
        strTmp = "���͡�" & strInfo & "�����ݳɹ�"
    Else
        '�쳣��־���
        strTmp = "���͡�" & strInfo & "������ʧ��"
    End If
    mclsLog.Add strTmp
    mclsLog.Save
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    Exit Sub
    
hErr:
    RaiseEvent AfterTransmit(Err.Description, mblnTransmitting)
    mclsLog.Add Err.Description
    mclsLog.Save
End Sub

Public Sub TransBase(ByVal strData As String)
'���ܣ��������ݴ���
'������
'  strData�����clsDataTransmit.Transmit
'  strData�����͵�����
'    ��ʽ��[�ӿڱ��]|[ҵ�����]|[ҵ������]
'    �ӿڱ�ţ��ӿڱ��1[;�ӿڱ��...]
'    ҵ�����
'        1-������Ϣ��
'        2-��Ա��Ϣ��
'        3-ҩƷĿ¼��
'        4-ҩƷ������λ��
'        5-��ҩ���ڣ�
'    ҵ�����ݣ�
'        1-���ţ���������1;��������2;��
'        2-��Ա����Ա����1;��Ա����2;��
'        3-ҩƷ������1;����2;��
'        4-��棺�ⷿid1;�ⷿid2;��
'        5-���ڣ��ⷿid1;�ⷿid2;��

    Dim strINF As String, strClass As String, strDetail As String, strTrans As String
    Dim strMsg As String, strTmp As String
    Dim arrItems As Variant, arrINF As Variant
    Dim i As Integer, j As Integer
    
    arrItems = Split(strData, "|")
    
    mclsLog.Add "" & Now, Val("1-��")
    mclsLog.Add "��ʼ�������ݴ���", Val("1-��")
    
    On Error GoTo hErr
    
    '�ӿڱ��
    arrINF = Split(arrItems(0), ";")
    For i = LBound(arrINF) To UBound(arrINF)
    
        strINF = Trim(arrINF(i))
        If strINF = "" Then GoTo Continue
        
        'ҵ������
        strClass = arrItems(1)
        If UBound(arrItems) > 1 Then
            strDetail = arrItems(2)
        Else
            strDetail = ""
        End If
        strTrans = strINF & "|" & strDetail
                
        Select Case Val(strClass)
        Case 1
            strTmp = "��" & strINF & "���ӿڴ��͡�������Ϣ��"
        Case 2
            strTmp = "��" & strINF & "���ӿڴ��͡���Ա��Ϣ��"
        Case 3
            strTmp = "��" & strINF & "���ӿڴ��͡�ҩƷĿ¼��"
        Case 4
            strTmp = "��" & strINF & "���ӿڴ��͡������Ϣ��"
        Case 5
            strTmp = "��" & strINF & "���ӿڴ��͡���ҩ���ڡ�"
        End Select
        
        '��������
        mclsLog.Add strTrans, Val("1-��"), Val("1-��ϸ��־")
        
        If mobjMachine.Operation(mstrUser, Val(strClass), strTrans, strMsg) Then
            '������־���
            strTmp = strTmp & "�ɹ�"
        Else
            '�쳣��־���
            strTmp = strTmp & "ʧ��"
        End If
        
        mclsLog.Add strTmp, Val("1-��")
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
Continue:
    Next
    
    Erase arrINF
    Erase arrItems
    
    Exit Sub
    
hErr:
    mclsLog.Add Err.Number & ":" & Err.Description, Val("1-��")
    mclsLog.Save
End Sub


