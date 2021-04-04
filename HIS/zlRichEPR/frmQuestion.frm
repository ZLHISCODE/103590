VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuestion 
   Caption         =   "���Ӳ����������"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmQuestion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   6555
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQuestion.frx":000C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8652
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsCondition As ADODB.Recordset
Private mfrmParent As Object
Private mblnOK As Boolean
Private mlngModul As Long
Private mstrPrivs As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnAuditEnter As Boolean '��������¼��������

Private WithEvents mfrmChildQuestion As frmChildQuestion
Attribute mfrmChildQuestion.VB_VarHelpID = -1
Public Event ShowInfo(ByVal strShowInfo As Long)


Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestion.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestion Is Nothing) Then
        DataChanged = mfrmChildQuestion.DataChanged
    End If
End Property

'################################################################################################################
'   ��;��  ϵͳ��ڡ�
'################################################################################################################
Public Sub ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long)

    On Error GoTo errHand
    Dim lng�ύId As Long
    Dim lng��Ժ����ID As Long
    '��ʼ��������
    Set mfrmParent = frmParent
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    
    '��ʼϵͳ����
    mlngModul = 1560
    mblnAuditEnter = True '��������¼��������
    mstrPrivs = GetPrivFunc(glngSys, mlngModul) '��ȡȨ��
    
    '��ʼ��
    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ʼ����")
    Call ExecuteCommand("�������")
    
    
    '��ʾ����
    
    'Me.Show vbModal, mfrmParent
    mfrmChildQuestion.Show vbModal, mfrmParent
    
    If mblnOK Then
        
    Else
 
    End If
    
    Set mrsCondition = Nothing
    If Not (mfrmChildQuestion Is Nothing) Then Unload mfrmChildQuestion
    
    Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop         As Integer
    Dim strTmp As String
    
    On Error GoTo errHand

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        Set mfrmChildQuestion = New frmChildQuestion
        Call mfrmChildQuestion.InitData(Me, mlngModul, IsPrivs(mstrPrivs, "��鲡��"), mblnAuditEnter, mstrPrivs)
   
        
     Case "��ʼ����"
                                
        '��������������Ŀ�������г�ʼ��
        Call ParamCreate(mrsCondition)
        
        Call ParamAdd(mrsCondition, "�ȴ�����", 1)
        Call ParamAdd(mrsCondition, "�ܾ�����", 1)
        Call ParamAdd(mrsCondition, "�������", 1)
        Call ParamAdd(mrsCondition, "��鷴��", 1)
        Call ParamAdd(mrsCondition, "�������", 1)
        
        Call ParamAdd(mrsCondition, "��ǰ����", "")
        Call ParamAdd(mrsCondition, "��Ժ���", "")
        
        Call ParamAdd(mrsCondition, "��������", 0)
        Call ParamAdd(mrsCondition, "ҽ������", "")
        
        Call ParamAdd(mrsCondition, "��鿪ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "������ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�鵵��ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�鵵����ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    
        Call ParamAdd(mrsCondition, "��Ժ����", 0)
        Call ParamAdd(mrsCondition, "��Ժ��ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "��Ժ����ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        
        Call ParamAdd(mrsCondition, "ҽ����ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "ҽ������ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "סԺҽʦ", "")
        Call ParamAdd(mrsCondition, "��������", "")
        Call ParamAdd(mrsCondition, "�������", "")
        Call ParamAdd(mrsCondition, "ҩƷ��Ϣ", "")
                
        '��ȡȱʡʱ�䷶Χ
        strTmp = GetPara("���ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "��鿪ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "������ʱ��", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("�鵵ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "�鵵��ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "�鵵����ʱ��", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("��Ժȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "��Ժ��ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "��Ժ����ʱ��", GetDateTime(strTmp, 2))
        
        '�¼�����
        strTmp = GetPara("ҽ��ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "ҽ����ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "ҽ������ʱ��", GetDateTime(strTmp, 2))
    Case "�������"
        Dim strObject As String
        Dim strParam As String
        Dim lng�ύId As Long
        
        If Not (mfrmChildQuestion Is Nothing) Then
            strObject = "��ҳ��¼"
            lng�ύId = GetSubmitID(mlng����ID, mlng��ҳID)
            Call mfrmChildQuestion.SetParamter(mlng����ID, mlng��ҳID, strObject, strParam, lng�ύId)
            mfrmChildQuestion.AllowModify = True
            Call mfrmChildQuestion.RefreshData("", mrsCondition, mblnAuditEnter) 'GetChildPatient(mintIndex).Depts
            
            
        End If
    
    End Select
    
    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:
End Function

Public Property Get ģ���() As Long
    ģ��� = mlngModul
End Property

Private Sub mfrmChildQuestion_AfterDataChanged()
    ' Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildQuestion_AfterDeleteQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    ' Call ExecuteCommand("ˢ��ָ������", lng����ID, lng��ҳID)
End Sub

Private Sub mfrmChildQuestion_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True Ժ������ =Flase �Ƽ�����
'    If blnQuestionType Then
'        If ObjPtr(dkpMain.Panes(1)) > 0 Then
'            dkpMain.Panes(1).Title = "Ժ�����ⷴ��"
'        End If
'    Else
'        If ObjPtr(dkpMain.Panes(1)) > 0 Then
'            dkpMain.Panes(1).Title = "�Ƽ����ⷴ��"
'        End If
'    End If
End Sub

Private Sub mfrmChildQuestion_AfterSaveQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    ' Call ExecuteCommand("ˢ��ָ������", lng����ID, lng��ҳID)
End Sub

Private Sub mfrmChildQuestion_LocationDocument(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt�������� As Byte, ByVal lng�ļ�ID As Long, ByVal lngҽ��id As Long, ByVal lng����ID As Long)
    '������Ϣ��λ��ָ�����˵�ָ������������ȥ
    On Error GoTo errHand
    
    
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub
