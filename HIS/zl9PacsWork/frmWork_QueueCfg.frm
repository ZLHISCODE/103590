VERSION 5.00
Begin VB.Form frmWork_QueueCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   Icon            =   "frmWork_QueueCfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkQueueQuick 
      Caption         =   "�Զ�������ݺ��д���"
      Height          =   180
      Left            =   240
      TabIndex        =   36
      Top             =   3960
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkLockAfterCall 
      Caption         =   "���к������ɼ�"
      Height          =   180
      Left            =   3240
      TabIndex        =   35
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox chkShowMySelfCalled 
      Caption         =   "ֻ��ʾ�Լ����еĶ���"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   3540
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ������(&P)"
      Height          =   375
      Left            =   1635
      Picture         =   "frmWork_QueueCfg.frx":1042
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Frame framColumn 
      Caption         =   "�Ŷ�������"
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   1005
      Width           =   7095
      Begin VB.CheckBox chkColumn 
         Caption         =   "ҽ������"
         Height          =   255
         Index           =   6
         Left            =   1305
         TabIndex        =   28
         Tag             =   "ҽ������"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "��ע"
         Height          =   255
         Index           =   9
         Left            =   5505
         TabIndex        =   27
         Tag             =   "��ע"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�ŶӺ���"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "�ŶӺ���"
         Top             =   375
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Tag             =   "��������"
         Top             =   375
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�Ա�"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   24
         Tag             =   "�Ա�"
         Top             =   375
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "����(ִ�м�)"
         Height          =   255
         Index           =   4
         Left            =   5505
         TabIndex        =   23
         Tag             =   "����"
         Top             =   375
         Width           =   1440
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�����Ŀ"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Tag             =   "�����Ŀ"
         Top             =   720
         Width           =   1065
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "��ǰ״̬"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   21
         Tag             =   "�Ŷ�״̬"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "�Ŷ�ʱ��"
         Height          =   255
         Index           =   8
         Left            =   4020
         TabIndex        =   20
         Tag             =   "�Ŷ�ʱ��"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   4020
         TabIndex        =   19
         Tag             =   "����"
         Top             =   375
         Width           =   1095
      End
   End
   Begin VB.Frame framCalledColumn 
      Caption         =   "����������"
      Height          =   1080
      Left            =   225
      TabIndex        =   10
      Top             =   2175
      Width           =   7110
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "��ע"
         Height          =   255
         Index           =   10
         Left            =   5790
         TabIndex        =   32
         Tag             =   "��ע"
         Top             =   705
         Width           =   705
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "��ǰ״̬"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   31
         Tag             =   "�Ŷ�״̬"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "ҽ������"
         Height          =   255
         Index           =   6
         Left            =   105
         TabIndex        =   30
         Tag             =   "ҽ������"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����(ִ�м�)"
         Height          =   255
         Index           =   4
         Left            =   4290
         TabIndex        =   29
         Tag             =   "����"
         Top             =   360
         Width           =   1425
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "�����Ŀ"
         Height          =   255
         Index           =   5
         Left            =   5805
         TabIndex        =   17
         Tag             =   "�����Ŀ"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����ʱ��"
         Height          =   255
         Index           =   8
         Left            =   2835
         TabIndex        =   16
         Tag             =   "����ʱ��"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "������"
         Height          =   255
         Index           =   7
         Left            =   1575
         TabIndex        =   15
         Tag             =   "����ҽ��"
         Top             =   705
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1305
         TabIndex        =   14
         Tag             =   "��������"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "�ŶӺ���"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "�ŶӺ���"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "�Ա�"
         Height          =   255
         Index           =   2
         Left            =   2610
         TabIndex        =   12
         Tag             =   "�Ա�"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   3435
         TabIndex        =   11
         Tag             =   "����"
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdVoiceCfg 
      Caption         =   "��������(&V)"
      Height          =   375
      Left            =   225
      Picture         =   "frmWork_QueueCfg.frx":118C
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   5115
      Picture         =   "frmWork_QueueCfg.frx":12D6
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   6210
      Picture         =   "frmWork_QueueCfg.frx":1420
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1100
   End
   Begin VB.ComboBox cbxTurnPage 
      Height          =   300
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3480
      Width           =   2580
   End
   Begin VB.Frame frmRoomCfg 
      Caption         =   "����ִ�м�����"
      Height          =   765
      Left            =   270
      TabIndex        =   0
      Top             =   165
      Width           =   7050
      Begin VB.ComboBox cbxRoomName 
         Height          =   300
         Left            =   4635
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2325
      End
      Begin VB.ComboBox cbxDept 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   2310
      End
      Begin VB.Label Label2 
         Caption         =   "ִ�м����ƣ�"
         Height          =   195
         Left            =   3555
         TabIndex        =   2
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "�������ң�"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.Label Label3 
      Caption         =   "�������תҳ�棺"
      Height          =   240
      Left            =   3225
      TabIndex        =   5
      Top             =   3540
      Width           =   1455
   End
End
Attribute VB_Name = "frmWork_QueueCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlngModule As Long
Private mobjQueue As Object
Private mstrPrivs As String
Private mblnLockAfterCall As Boolean
Private mblnQueueQucik As Boolean


Public Function ShowQueueConfig(objQueue As Object, _
                                ByVal lngModule As Long, _
                                ByVal strPrivs As String, _
                                Optional objOwner As Object = Nothing, _
                                Optional ByRef blnLockAfterCall As Boolean = False, _
                                Optional ByRef blnQueueQuick As Boolean = False) As Boolean
'��ʾpacs��������
    ShowQueueConfig = False
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    Set mobjQueue = objQueue
    
    Call LoadTurnPage
    Call LoadStudyDept
    Call ReadCfgParameter
    
    CheckAddHeight
    Me.Show 1, objOwner

    blnLockAfterCall = mblnLockAfterCall
    ShowQueueConfig = mblnOK
    blnQueueQuick = mblnQueueQucik
End Function


Private Sub ReadCfgParameter()
'��ȡ���ò���
    Dim i As Long
    Dim strColumnInfo As String
    
    If mlngModule = 1291 Then
        chkLockAfterCall.value = zlDatabase.GetPara("���к������ɼ�", glngSys, mlngModule, "0")
        mblnLockAfterCall = chkLockAfterCall.value
    End If
    
    '��ȡ�ŶӶ�����Ϣ����
    strColumnInfo = zlDatabase.GetPara("�ŶӶ�����Ϣ����", glngSys, mlngModule, "�ŶӺ���,��������")
    
    For i = 0 To 9
        chkColumn(i).value = Int(IIf(InStr(1, "," & strColumnInfo & ",", "," & chkColumn(i).tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    '��ȡ���ж�����Ϣ����
    strColumnInfo = zlDatabase.GetPara("���ж�����Ϣ����", glngSys, mlngModule, "�ŶӺ���,��������")
    
    For i = 0 To 9
        chkCalledColumn(i).value = Int(IIf(InStr(1, "," & strColumnInfo & ",", "," & chkCalledColumn(i).tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    chkShowMySelfCalled.value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\�Ŷӽк�", "ֻ��ʾ�Լ����еĶ���", "1"))
    
    chkQueueQuick.value = Val(zlDatabase.GetPara("�Զ�������ݺ��д���", glngSys, mlngModule, "1"))
    mblnQueueQucik = IIf(chkQueueQuick.value = 1, True, False)
End Sub



Private Sub cbxDept_Click()
On Error GoTo errHandle
    Call LoadExeRoom(cbxDept.ItemData(cbxDept.ListIndex))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkLockAfterCall_Click()
    mblnLockAfterCall = chkLockAfterCall.value
End Sub


Private Sub chkQueueQuick_Click()
On Error GoTo errHandle
    mblnQueueQucik = IIf(chkQueueQuick.value = 1, True, False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    
    Unload Me
End Sub


Private Sub cmdPrintSet_Click()
'���ô�ӡ��
On Error GoTo errHandle
    Call mobjQueue.QueueOper.PrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
'�������
On Error GoTo errHandle
    Dim i As Long
    Dim strColumnInf As String
    
    zlDatabase.SetPara "����ִ�м����", cbxDept.ItemData(cbxDept.ListIndex), glngSys, mlngModule
    zlDatabase.SetPara "����ִ�м�����", cbxRoomName.Text, glngSys, mlngModule
    zlDatabase.SetPara "�������תҳ��", cbxTurnPage.Text, glngSys, mlngModule
    If mlngModule = 1291 Then zlDatabase.SetPara "���к������ɼ�", chkLockAfterCall.value, glngSys, mlngModule
    
    '�����ŶӶ�������
    strColumnInf = ""
    For i = 0 To 9
        If chkColumn(i).value = vbChecked Or chkColumn(i).tag = "�ŶӺ���" Or chkColumn(i).tag = "��������" Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkColumn(i).tag
        End If
    Next i
    
    zlDatabase.SetPara "�ŶӶ�����Ϣ����", strColumnInf, glngSys, mlngModule
    
    '������ж�������
    strColumnInf = ""
    For i = 0 To 10
        If chkCalledColumn(i).value = vbChecked Or chkCalledColumn(i).tag = "�ŶӺ���" Or chkCalledColumn(i).tag = "��������" Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkCalledColumn(i).tag
        End If
    Next i
    
    zlDatabase.SetPara "���ж�����Ϣ����", strColumnInf, glngSys, mlngModule
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\�Ŷӽк�", "ֻ��ʾ�Լ����еĶ���", chkShowMySelfCalled.value
    
    zlDatabase.SetPara "�Զ�������ݺ��д���", chkQueueQuick.value, glngSys, mlngModule
    
    mblnOK = True
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdVoiceCfg_Click()
'���������ô���
On Error GoTo errHandle
    If mobjQueue Is Nothing Then Exit Sub
    
    Call mobjQueue.ShowVoiceConfig

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadTurnPage()
'�����������תҳ��
    Dim strTurnPage As String
    Dim strPages() As String
    Dim i As Long
    
    strTurnPage = zlDatabase.GetPara("�������תҳ��", glngSys, mlngModule, "")

    If mlngModule = 1290 Then
        strPages = Split("Ӱ��,����,����,ҽ��,����", ",")
    Else
        strPages = Split("�ɼ�,����,����,ҽ��,����", ",")
    End If
    
    cbxTurnPage.Clear
    
    Call cbxTurnPage.AddItem("")
    
    For i = 0 To UBound(strPages)
        If Trim(strPages(i)) <> "" Then
            cbxTurnPage.AddItem strPages(i)
        End If
        
        If Trim(strPages(i)) = strTurnPage Then
            cbxTurnPage.ListIndex = i + 1
        End If
    Next i
    
    If cbxTurnPage.ListIndex < 0 Then cbxTurnPage.ListIndex = 0
End Sub


Private Sub LoadStudyDept()
'���������
    Dim strSql As String
    Dim rsData As New ADODB.Recordset
    Dim str��Դ As String
    Dim strCfgDept As String
    
    str��Դ = "1,2,3"
    
    strCfgDept = zlDatabase.GetPara("����ִ�м����", glngSys, mlngModule)

    If CheckPopedom(mstrPrivs, "���п���") Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If

    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯִ�м�����������", CStr("," & str��Դ & ","))
    

    Do Until rsData.EOF
        cbxDept.AddItem Nvl(rsData!����)
        cbxDept.ItemData(cbxDept.ListCount - 1) = Val(Nvl(rsData!ID))
        
        If Nvl(rsData!ID) = strCfgDept Then
            cbxDept.ListIndex = cbxDept.ListCount - 1
        End If
        
        rsData.MoveNext
    Loop
        
    If cbxDept.ListCount > 0 And cbxDept.ListIndex < 0 Then cbxDept.ListIndex = 0
End Sub

Private Sub LoadExeRoom(ByVal lngDeptID As Long)
'����ִ�м�
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCfgRoom As String
    
    strCfgRoom = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule)
    
    strSql = "select ִ�м�,�豸�� from ҽ��ִ�з��� a, Ӱ���豸Ŀ¼ b Where a.����豸=b.�豸��(+) and ����ID=[1] order by ִ�м�"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ����ִ�м�", lngDeptID)
    
    cbxRoomName.Clear
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        cbxRoomName.AddItem Nvl(rsData!ִ�м�) & "-" & Nvl(rsData!�豸��)
        
        If Nvl(rsData!ִ�м�) & "-" & Nvl(rsData!�豸��) = strCfgRoom Then
            cbxRoomName.ListIndex = cbxRoomName.ListCount - 1
        End If
        
        rsData.MoveNext
    Wend
    
    If cbxRoomName.ListCount > 0 And cbxRoomName.ListIndex <= 0 Then cbxRoomName.ListIndex = 0
End Sub

Private Sub CheckAddHeight()
    '�ж��Ƿ���Ҫ���Ӵ���߶ȣ�104686��أ�����ǲɼ�����վ����Ҫ���Ӹ߶�����ʾ�����к������ɼ����������
    If mlngModule = 1291 Then
        chkLockAfterCall.Visible = True
    Else
        chkLockAfterCall.Visible = False
    End If
End Sub


