VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIS�����ϴ� v1.0"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7110
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdStore 
      Caption         =   "��ȡ�豸ҩƷ���"
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdDept 
      Caption         =   "���������ϴ�(&D)"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Tag             =   "0"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   5880
      TabIndex        =   3
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Timer TimerTrans 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   6885
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "��ʼ�ϴ�(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox lstLog 
      Height          =   5280
      ItemData        =   "frmMain.frx":030A
      Left            =   120
      List            =   "frmMain.frx":030C
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng��ѯ��� As Long
Private mint��ѯ���� As Integer
Private mstr��ʼʱ�� As String
Private mstr����ʱ�� As String
Private mstrҩ��id As String
Private mblnPackerConnect As Boolean
Private mblnExit As Boolean
Private mstrUserCode As String
Private mstrUserName As String

Private Sub AutoTrans()
    Dim rsData As ADODB.Recordset
    Dim strUserCode As String
    Dim strUserName As String
    Dim strReturn As String
    
    On Error GoTo errHandle
    
    '�������ڷ�Χ
    Call UpdateDateValue
    
    Me.cmdStart.Enabled = False
    
    Call OutputLog("��ʼ��ȡ������Ϣ")
        
    '��NO�����ϴ����ݣ�����סԺ���ݣ������õ���=8�����շ�=1���֣�סԺ�õ���=9���֣�
    gstrSql = "Select ����, NO " & vbNewLine & _
        " From δ��ҩƷ��¼ " & vbNewLine & _
        " Where (���� = 8 And ���շ� = 1 Or ���� = 9) And Nvl(�Ƿ��ϴ�, 0) = 0 And �������� Between [1] And [2] And " & vbNewLine & _
        " �ⷿid In (Select * From Table(Cast(f_Num2list([3], ';') As Zltools.t_Numlist))) "
    Set rsData = OpenSQLRecord(gstrSql, "AutoTrans", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), mstrҩ��id)
    
    If Not gobjPacker Is Nothing And mblnPackerConnect = True Then
        If rsData.EOF = False Then
            Do While Not rsData.EOF
                Call gobjPacker.DYEY_MZ_TransRecipeDetail(1, mstrUserCode, mstrUserName, 0, rsData!���� & "," & rsData!NO, strReturn)
                
                LogListItem "�����ϴ��ɹ���" & rsData!NO
                Call OutputLog("�����ϴ��ɹ���" & rsData!NO)
                rsData.MoveNext
            Loop
            LogListItem "�����ϴ�������ɣ�" & Now
            Call OutputLog("�����ϴ�������ɣ�" & Now)
        Else
            LogListItem "�����������ϴ���" & Now
            Call OutputLog("�����������ϣ�" & Now)
        End If
    Else
        LogListItem "WebService��ַ����ȷ��" & Now
        Call OutputLog("WebService��ַ����ȷ��" & Now)
    End If
  
    Me.cmdStart.Enabled = True
    
    Exit Sub
    
errHandle:
    Me.cmdStart.Enabled = True
    LogListItem Err.Description
End Sub

Private Sub cmdDept_Click()
    Dim strMsg As String
    
    If gobjPacker Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    cmdDept.Enabled = False
    If gobjPacker.DYEY_MZ_TransDept("", mstrUserCode, mstrUserName, strMsg) Then
        LogListItem "���������ϴ��ɹ���" & Now
    Else
        LogListItem "���������ϴ�ʧ�ܣ���鿴��־�ļ�ȷ��ԭ��" & Now
    End If
    cmdDept.Enabled = True
    Exit Sub
    
hErr:
    cmdDept.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim strMsg As String
    
    If cmdStart.Tag = "0" Then
        If mblnPackerConnect = False Then
            '���³�ʼ��
            TimerTrans.Enabled = False
            mblnPackerConnect = gobjPacker.DYEY_MZ_IniSoap(False, strMsg, , gcnOracle, 1)
            If mblnPackerConnect = False Then
                MsgBox "��ʼ���ӿڲ���ʧ�ܣ�Soap��ʼ��ʧ�ܣ�", vbInformation, "��ʾ��Ϣ"
                Call OutputLog("��ʼ���ӿڲ���ʧ�ܣ�Soap��ʼ��ʧ��")
                Exit Sub
            End If
        End If
                
        cmdStart.Tag = "1"
        cmdStart.Caption = "ֹͣ�ϴ�(&S)"
        
        '��ʼ�ϴ�
        TimerTrans.Enabled = True
        
        LogListItem "��ʼ�ϴ���" & Now
        
    Else

        cmdStart.Tag = "0"
        cmdStart.Caption = "��ʼ�ϴ�(&S)"
        
        'ֹͣ�ϴ�
        TimerTrans.Enabled = False
        
        LogListItem "ֹͣ�ϴ�" & Now
        
    End If
    
    cmdDept.Enabled = cmdStart.Tag = "0"

End Sub

Private Sub cmdStore_Click()
    Dim strStore As String
    
    cmdStore.Enabled = False
    Call ReadDeviceStore(strStore)
    Call WriteZLHIS(strStore)
    cmdStore.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnExit Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    
    '��ȡע������
    mlng��ѯ��� = Val(GetSetting("ZLSOFT", "����ģ��\�Զ���ҩ��", "��ѯ���", 60))
    mint��ѯ���� = Val(GetSetting("ZLSOFT", "����ģ��\�Զ���ҩ��", "��ѯ����", 0))
    mstrҩ��id = GetSetting("ZLSOFT", "����ģ��\�Զ���ҩ��", "����ҩ��", "")
    mblnExit = False
    
    If mstrҩ��id = "" Then
        MsgBox "δע��ҩ����Ϣ����ʼ��ʧ�ܣ�", vbInformation, ""
        Call OutputLog("δע��ҩ����Ϣ����ʼ��ʧ�ܣ�")
        Unload Me
    End If
    
    If mlng��ѯ��� > 60 Then
        mlng��ѯ��� = 60
    End If
    TimerTrans.Interval = mlng��ѯ��� * 1000
    
    'ȡ�û���Ϣ
    Call GetUserInfo
    
    '�������ڷ�Χ
    Call UpdateDateValue
    
    '�Զ���ҩ���ӿ�
    On Error Resume Next
    Set gobjPacker = CreateObject("zlDrugPacker.clsDrugPacker")
    Err.Clear
    If gobjPacker Is Nothing Then
        MsgBox "��ʼ���Զ���ҩ���ӿڲ���ʧ�ܣ�", vbInformation, "��ʾ��Ϣ"
        Call OutputLog("��ʼ���ӿڲ���ʧ�ܣ������ӿڲ���ʧ��")
        mblnExit = True
        Exit Sub
    Else
        mblnPackerConnect = gobjPacker.DYEY_MZ_IniSoap(False, strMsg, , gcnOracle, 1)
        If mblnPackerConnect = False Then
            MsgBox "��ʼ���ӿڲ���ʧ�ܣ�Soap��ʼ��ʧ�ܣ�", vbInformation, "��ʾ��Ϣ"
            Call OutputLog("��ʼ���ӿڲ���ʧ�ܣ�Soap��ʼ��ʧ��")
            Exit Sub
        End If
    End If
    
    Call OutputLog("��ȡ������" & "��ѯ���=" & mlng��ѯ��� & "," & "��ѯ����=" & mint��ѯ���� & "," & "ҩ��ID=" & mstrҩ��id)
    Call OutputLog("��ʼ���ӿڲ����ɹ�")
End Sub

Private Sub UpdateDateValue()
    If mint��ѯ���� = 0 Then
        'Ĭ���ǵ���
        mstr��ʼʱ�� = Format(Currentdate, "YYYY-MM-DD")
        mstr����ʱ�� = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    Else
        'ָ��������
        If mint��ѯ���� > 3 Then
            mint��ѯ���� = 3
        End If
        
        mstr��ʼʱ�� = Format(DateAdd("d", -mint��ѯ����, Currentdate), "YYYY-MM-DD")
        mstr����ʱ�� = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStart.Enabled = False Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set gobjComLib = Nothing
End Sub

Private Sub TimerTrans_Timer()
    TimerTrans.Enabled = False
    
    On Error GoTo errHandle
    
    '�������
    If gcnOracle.State <> adStateOpen Then
        gcnOracle.Open
    End If
    
    '�����Զ��ϴ�����
    DoEvents
    Call AutoTrans
    DoEvents
    
    TimerTrans.Enabled = True
    
    Exit Sub
    
errHandle:
    Call LogListItem("�쳣��" & Err.Description)
    TimerTrans.Enabled = True
End Sub

Private Sub LogListItem(ByVal strLog As String)
    Const INT_MAX_LINES As Integer = 200

    Me.lstLog.AddItem strLog
    Me.lstLog.Selected(Me.lstLog.ListCount - 1) = True
    Me.lstLog.TopIndex = Me.lstLog.ListCount - 1
    If lstLog.ListCount >= INT_MAX_LINES Then lstLog.RemoveItem 0

End Sub

Private Sub GetUserInfo()
'ȡ�û���Ϣ

    Dim rsData As ADODB.Recordset

    On Error GoTo hErr
    gstrSql = "Select a.���, a.����, b.�û��� From ��Ա�� A, �ϻ���Ա�� B Where a.Id = b.��Աid And Upper(�û���) = Upper([1])"
    Set rsData = OpenSQLRecord(gstrSql, "�û���Ϣ", gstrUser)
    If rsData.RecordCount > 0 Then
        mstrUserCode = rsData!���
        mstrUserName = rsData!����
    Else
        mstrUserCode = ""
        mstrUserName = ""
    End If
    rsData.Close
    Exit Sub
    
hErr:
    mstrUserCode = ""
    mstrUserName = ""
    MsgBox Err.Description, vbInformation, "��ʾ"
End Sub

Private Sub ReadDeviceStore(ByRef strVar As String)
'��ȡ�豸�Ŀ��

    Dim strMsg As String

    strVar = ""
    If gobjPacker Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    '��ȡ�豸���
    If gobjPacker.DYEY_MZ_TransStockDevice(mstrUserCode, mstrUserName, strVar, strMsg) Then
        LogListItem "�豸����������سɹ���" & Now
    Else
        LogListItem "�豸�����������ʧ�ܣ���鿴��־�ļ�ȷ��ԭ��" & Now
        LogListItem strMsg
    End If
    
    Exit Sub
    
hErr:
    MsgBox Err.Description, vbInformation, "��ʾ��Ϣ"
End Sub

Private Sub WriteZLHIS(ByVal strVar As String)
'д��ZLHIS���ݱ�
    
    Dim strSQL As String, strTmp As String
    Dim intReturn As Integer
    Dim strDisp As String, strCode As String
    Dim dblQTY As Double
    Dim cmdInsert As ADODB.Command
    
'<ROOT>
'    <RETCODE>1</RETCODE>
'    <CONSIS_DRUG_BATCHVW>
'        <DISPENSARY>320</DISPENSARY>
'        <DRUG_CODE>1-1009</DRUG_CODE>
'        <QUANTITY>9</QUANTITY>
'    </CONSIS_DRUG_BATCHVW>
'    <CONSIS_DRUG_BATCHVW>
'        <DISPENSARY>86</DISPENSARY>
'        <DRUG_CODE>1-1015</DRUG_CODE>
'        <QUANTITY>12</QUANTITY>
'    </CONSIS_DRUG_BATCHVW>
'</ROOT>
    
    If strVar = "" Then Exit Sub
    
    '����XML
    If InStr(strVar, "<RETCODE>") <= 0 Then Exit Sub
    
    intReturn = Val(Mid(strVar, InStr(strVar, "<RETCODE>") + 9))
    If intReturn = 1 Then
        '���óɹ�
        
        On Error GoTo hErr
        
        ''������ݱ�
        gcnOracle.Execute "Delete DrugDeviceStoreTemp"
        
        ''��д
        Do While InStr(strVar, "<CONSIS_DRUG_BATCHVW>") > 0
            strVar = Mid(strVar, InStr(strVar, "<CONSIS_DRUG_BATCHVW>") + 29)
            
            'DISPENSARY
            If InStr(strVar, "<DISPENSARY>") > 0 Then
                strDisp = Mid(strVar, InStr(strVar, "<DISPENSARY>") + 12)
                strDisp = Left(strDisp, InStr(strDisp, "</") - 1)
            Else
                strDisp = ""
            End If
            
            'DRUG_CODE
            If InStr(strVar, "<DRUG_CODE>") > 0 Then
                strCode = Mid(strVar, InStr(strVar, "<DRUG_CODE>") + 11)
                strCode = Left(strCode, InStr(strCode, "</") - 1)
            Else
                strCode = ""
            End If
            
            'QUANTITY
            If InStr(strVar, "<QUANTITY>") > 0 Then
                dblQTY = Val(Mid(strVar, InStr(strVar, "<QUANTITY>") + 10))
            Else
                dblQTY = 0
            End If
                        
            'дZLHIS���ݱ�
            If strCode <> "" Then
                strSQL = "insert into DrugDeviceStoreTemp (�ⷿID,ҩƷ����,�������) values "
                strSQL = strSQL & "(" & IIf(strDisp = "", "null", strDisp) & ","
                strSQL = strSQL & "'" & strCode & "',"
                strSQL = strSQL & dblQTY & ")"
                
                Set cmdInsert = New ADODB.Command
                With cmdInsert
                    .ActiveConnection = gcnOracle
                    .CommandText = strSQL
                    .Execute
                End With
            Else
                Debug.Print "��ҩƷ���룡"
            End If
        Loop
    ElseIf intReturn = 0 Then
        LogListItem "�豸�޿�����ݣ�" & Now
    Else
        LogListItem "�豸������������쳣��WebService����" & Now
    End If
    
    Exit Sub
    
hErr:
    MsgBox Err.Description, vbInformation, "��ʾ��Ϣ"
End Sub
