VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendBody 
   Caption         =   "������ͼ"
   ClientHeight    =   7350
   ClientLeft      =   180
   ClientTop       =   450
   ClientWidth     =   10740
   Icon            =   "frmCaseTendBody.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCustom 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2265
      ScaleHeight     =   300
      ScaleWidth      =   1965
      TabIndex        =   2
      Top             =   5220
      Width           =   1965
      Begin VB.CommandButton cmd 
         Height          =   300
         Left            =   1665
         Picture         =   "frmCaseTendBody.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1665
      End
   End
   Begin zl9BodyEditorSXHZ.usrBodyEditor BodyEdit 
      Height          =   4350
      Left            =   435
      TabIndex        =   0
      Top             =   615
      Width           =   6375
      _extentx        =   11245
      _extenty        =   7673
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6990
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":6AD8
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'�ֲ�������������

Private mrsParam As New ADODB.Recordset
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mblnChildForm As Boolean
Private mblnOK As Boolean
Private mfrmMain As Object
Private mblnChanged As Boolean
Private mcbr�鿴 As CommandBarControl
Private mstr���²�λ As String
Private mstr������ʽ As String
Private mstr���� As String
Private mcbrMenuBar���� As CommandBarControl
Private mcbrMenuBar��λ As CommandBarControl
Private mcbrMenuBar�༭ As CommandBarControl
Private mcbrToolBar As CommandBar
Private mint���￨���볤�� As Integer
Private mstrSvr���� As String
Private mrsPatient As ADODB.Recordset
Private mlngRowNum As Long
Private mstrFindKey As String
Private mobjFindKey As CommandBarControl
Private mstrPrivs As String
Private mblnShowing As Boolean

Public Event AfterPrint()

'######################################################################################################################
'�Զ��庯������������

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mblnStartUp = True
    mblnChanged = False
    mstrPrivs = strPrivs
    mstr���²�λ = "Ҹ��"
    mstr������ʽ = "��������"
    mstr���� = ""
    
    blnShowing = mblnShowing
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    If blnShowing Then
        If Val(varParam(0)) = Val(mrsParam("����id").Value) Or Val(varParam(1)) = Val(mrsParam("��ҳid").Value) And Val(mrsParam("����id").Value) = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set mfrmMain = frmMain

    '������ʽ������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
    
    '��ʼ������
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "��ҳid", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "��Ժ", adTinyInt)
    Call CreateParam(mrsParam, "Ӥ��", adTinyInt)
    Call CreateParam(mrsParam, "�༭", adTinyInt)
    Call CreateParam(mrsParam, "����ȼ�", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ��ʼ����", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ��������", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "����Ʋ���", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)

    mrsParam.Open
    mrsParam.AddNew
    
    mrsParam("Ӥ��").Value = 0
    mrsParam("��Ժ").Value = 0
    mrsParam("�༭").Value = 0
    
    mrsParam("����id").Value = Val(varParam(0))
    mrsParam("��ҳid").Value = Val(varParam(1))
    mrsParam("����id").Value = Val(varParam(2))
    mrsParam("����id").Value = Val(varParam(2))
    
    If UBound(varParam) >= 3 Then mrsParam("��Ժ").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("�༭").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("Ӥ��").Value = Val(varParam(5))
    
    
    '��Ժ��ʼ����;��Ժ��������;��Ժ����;��Ժ����;ת������;ת������
    '------------------------------------------------------------------------------------------------------------------
    strPar = zlDatabase.GetPara("������ʾ��Χ", glngSys, 1262, "10000")
    mrsParam("��Ժ����").Value = Val(Mid(strPar, 1, 1))
    mrsParam("��Ժ����").Value = Val(Mid(strPar, 2, 1))
    mrsParam("ת������").Value = Val(Mid(strPar, 4, 1))
    On Error Resume Next
    mrsParam("����Ʋ���").Value = Val(Mid(strPar, 5, 1))
    On Error GoTo 0
    
    mrsParam("ת������").Value = Val(zlDatabase.GetPara("���ת������", 7))
    
    Dim curDate As Date
    Dim intDay As Integer
    
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, 1262, 7))
    mrsParam("��Ժ��������").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, 1262, 30))
    mrsParam("��Ժ��ʼ����").Value = Format(CDate(mrsParam("��Ժ��������").Value) - intDay, "yyyy-MM-dd 00:00:00")
    
    If blnShowing = False Then Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
    If rs.BOF = False Then
        mrsParam("����id").Value = Val(zlCommFun.NVL(rs("��Ժ����ID").Value))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ���� from ������Ϣ Where ����id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value))
    If rs.BOF = False Then
        txt.Text = zlCommFun.NVL(rs("����").Value)
        txt.Tag = ""
    End If
    
    '���￨����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ���ų��� from ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "���￨")
    If rs.BOF = False Then
        mint���￨���볤�� = Val(zlCommFun.NVL(rs("���ų���").Value))
    Else
        mint���￨���볤�� = 7
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If ReadPatient = False Then
        mblnStartUp = False
        Unload Me
        Exit Function
    End If
    
    mrsPatient.Filter = ""
    mrsPatient.Filter = "����id=" & Val(mrsParam("����id").Value)
    If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
    mrsPatient.Filter = ""
    
    Set BodyEdit.ParentForm = Me
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        mblnStartUp = False
        Unload Me
        Exit Function
    End If
    
    mblnStartUp = False
    
    If blnShowing = False Then
        Hook Me.hWnd
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        
        ShowEdit = mblnChanged
    End If
    
End Function

Public Function zlInit() As Boolean

    mblnChildForm = True

'    Call InitMenuBar

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '���:1-Ԥ��,2-��ӡ
    '����ֵ:0-ʧ��;1-�ɹ�;2-��ӡ
    gblnPrinted = False
    
'    If bytMode = 1 Then
'        zlPrintBody = PrintData(2, strPrintDevice)
'    Else
'        zlPrintBody = PrintData(1, strPrintDevice)
'    End If
    
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Public Function zlRefresh(strParam As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    picCustom.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    mblnStartUp = True
    mblnChanged = False
'    mstrPrivs = strPrivs
    mstr���²�λ = "Ҹ��"
    mstr������ʽ = "��������"
    mstr���� = ""
    
'    Set mfrmMain = frmMain
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    '������ʽ������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
    
    '��ʼ������
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "��ҳid", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "��Ժ", adTinyInt)
    Call CreateParam(mrsParam, "Ӥ��", adTinyInt)
    Call CreateParam(mrsParam, "�༭", adTinyInt)
    Call CreateParam(mrsParam, "����ȼ�", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ��ʼ����", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ��������", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "����Ʋ���", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)
    
    mrsParam.Open
    mrsParam.AddNew
    
    mrsParam("Ӥ��").Value = 0
    mrsParam("��Ժ").Value = 0
    mrsParam("�༭").Value = 0
    
    mrsParam("����id").Value = Val(varParam(0))
    mrsParam("��ҳid").Value = Val(varParam(1))
    mrsParam("����id").Value = Val(varParam(2))
    mrsParam("����id").Value = Val(varParam(2))
    
    If UBound(varParam) >= 3 Then mrsParam("��Ժ").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("�༭").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("Ӥ��").Value = Val(varParam(5))
    
    
    '��Ժ��ʼ����;��Ժ��������;��Ժ����;��Ժ����;ת������;ת������
    '------------------------------------------------------------------------------------------------------------------
    strPar = zlDatabase.GetPara("������ʾ��Χ", glngSys, 1262, "10000")
    mrsParam("��Ժ����").Value = Val(Mid(strPar, 1, 1))
    mrsParam("��Ժ����").Value = Val(Mid(strPar, 2, 1))
    mrsParam("ת������").Value = Val(Mid(strPar, 4, 1))
    On Error Resume Next
    mrsParam("����Ʋ���").Value = Val(Mid(strPar, 5, 1))
    On Error GoTo 0
    
    mrsParam("ת������").Value = Val(zlDatabase.GetPara("���ת������", 7))
    
    Dim curDate As Date
    Dim intDay As Integer
    
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, 1262, 7))
    mrsParam("��Ժ��������").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, 1262, 30))
    mrsParam("��Ժ��ʼ����").Value = Format(CDate(mrsParam("��Ժ��������").Value) - intDay, "yyyy-MM-dd 00:00:00")

'    Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
    If rs.BOF = False Then
        mrsParam("����id").Value = Val(zlCommFun.NVL(rs("��Ժ����ID").Value))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ���� from ������Ϣ Where ����id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value))
    If rs.BOF = False Then
        txt.Text = zlCommFun.NVL(rs("����").Value)
        txt.Tag = ""
    End If
    
    '���￨����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ���ų��� from ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "���￨")
    If rs.BOF = False Then
        mint���￨���볤�� = Val(zlCommFun.NVL(rs("���ų���").Value))
    Else
        mint���￨���볤�� = 7
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If ReadPatient = False Then
        mblnStartUp = False
'        Unload Me
        Exit Function
    End If
    
    mrsPatient.Filter = ""
    mrsPatient.Filter = "����id=" & Val(mrsParam("����id").Value)
    If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
    mrsPatient.Filter = ""
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        mblnStartUp = False
'        Unload Me
        Exit Function
    End If
    
    mblnStartUp = False
    
'    Hook Me.hWnd
        
    zlRefresh = True
    
End Function

Private Function ShowTxtSelDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rs As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional strInitKey As String = "", _
                                    Optional ByVal WinStyle As Byte = 3, _
                                    Optional ByVal blnSort As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:������+�б�ṹ
    '����:������2;�ɹ�����1;ȡ������0
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
        
    
    On Error GoTo errHand
    
    If rs.BOF Then Exit Function
    
    Call ClientToScreen(objTXT.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, WinStyle, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTXT.Height, strInitKey, strSavePath, , False, blnMuliSel, , blnSort) Then
                            
        Set rsResult = rs
        ShowTxtSelDialog = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strParam As String
    
    mstrSvr���� = txt.Text
    
    mrsParam("����ȼ�").Value = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rs.BOF = False Then mrsParam("����ȼ�").Value = zlCommFun.NVL(rs("����ȼ�"), 3)
    
    '��ʼ�����߲˵�
    If InitBodyLine = False Then Exit Function
    
    '����������ID,��ҳID,����ID,����ID,��Ժ��־;�༭��־;Ӥ��
    strParam = Val(mrsParam("����id")) & ";" & Val(mrsParam("��ҳid")) & ";" & Val(mrsParam("����id")) & ";" & Val(mrsParam("��Ժ")) & ";" & Val(mrsParam("�༭").Value) & ";" & Val(mrsParam("Ӥ��").Value)
    If Not BodyEdit.zlMenuClick("��ʼ����", strParam) Then Exit Function
'    If InitBody(Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("����id"))) = False Then Exit Function
        
    OpenPatientMap = True
    
End Function

Private Function ReadPatient() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strParam As String
    
    '��Ժ�ͳ�Ժ����:��Ժ���˿������ж��סԺ
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("��Ժ����").Value) <> 0 Or Val(mrsParam("��Ժ����").Value) <> 0 Or Val(mrsParam("����Ʋ���").Value) <> 0 Then
        gstrSQL = _
            "Select Decode(B.��Ժ����,NULL,Decode(B.״̬,3,2,1),Decode(B.��Ժ��ʽ,'����',4,3)) as ����," & _
            " Decode(B.��Ժ����,NULL,Decode(B.״̬,3,'Ԥ��Ժ����','��Ժ����'),Decode(B.��Ժ��ʽ,'����','��������','��Ժ����')) as ����," & _
            " A.����ID,B.��ҳID,B.סԺ��,A.�����,A.����,A.�Ա�,A.����,C.���� as ����,B.סԺҽʦ," & _
            " B.��Ժ���� as ����,B.�ѱ�,B.��Ժ����,B.��Ժ����,B.״̬,B.����,A.���￨��" & _
            " From ������Ϣ A,������ҳ B,���ű� C" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And ([6]=1 Or Nvl(B.״̬,0)<>1) And B.��Ժ����ID=C.ID" & _
            " And B.��ǰ����ID=[1] And ([4]<>0 And B.��Ժ���� is NULL Or [5]<>0 And B.��Ժ���� Between [2] And [3]) "
    End If
    
    'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("ת������").Value) <> 0 Then
        gstrSQL = gstrSQL & IIf(gstrSQL <> "", " Union All ", "") & _
            "Select Distinct 5 as ����,'ת������' as ����," & _
            " A.����ID,B.��ҳID,B.סԺ��,A.�����,A.����,A.�Ա�,A.����,D.���� as ����,C.����ҽʦ as סԺҽʦ," & _
            " C.����,B.�ѱ�,B.��Ժ����,B.��Ժ����,B.״̬,B.����,A.���￨��" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And C.����ID=D.ID" & _
            " And Nvl(B.״̬,0)=0 And B.��Ժ���� is NULL And B.��ǰ����ID<>[1]" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID=[1]" & _
            " And C.��ֹԭ��=3 And C.��ֹʱ�� Between Sysdate-[7] And Sysdate "
    End If
    gstrSQL = gstrSQL & " Order by ����,����,��ҳID Desc"
    gstrSQL = "Select RowNum As ID,1 As ĩ��,A.* From (" & gstrSQL & ") A"
    
    If Val(mrsParam("�༭").Value) = 1 Then
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("����id").Value), _
                                                                CDate(Format(mrsParam("��Ժ��ʼ����").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("��Ժ��������").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                0, _
                                                                Val(mrsParam("����Ʋ���").Value), _
                                                                Val(mrsParam("ת������").Value))
    Else
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("����id").Value), _
                                                                CDate(Format(mrsParam("��Ժ��ʼ����").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("��Ժ��������").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                Val(mrsParam("����Ʋ���").Value), _
                                                                Val(mrsParam("ת������").Value))
    End If
    
    ReadPatient = True
    
End Function

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    
    '�����˴�ӡ������,˵����������ӡ,�Զ��ӵ�1ҳ��ʼ��ӡ,�������κ�ѯ��
    '����:0-ȡ��,2-Ԥ��,1-��ӡ
    
    frmCaseTendBodyPrintSet.cmdPrint.Visible = (bytMode = 1)
    frmCaseTendBodyPrintSet.cmdPreview.Visible = (bytMode = 2)
    
    If strPrintDevice = "" Then
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, intPrintRange, lngBeginY, intBeginPage, mstrPrivs)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
            
    Select Case bytMode
    Case 2  '��ӡ
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice)
    Case 1  'Ԥ��
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice)
    End Select

    
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    
    On Error GoTo errHand
    
    If mcbrMenuBar���� Is Nothing Then
        InitBodyLine = True
        Exit Function
    End If
    
    mstrSQL = "SELECT A.��¼��,A.��Ŀ��� FROM ���¼�¼��Ŀ A,�����¼��Ŀ B " & _
            "WHERE A.��¼�� =1 And A.��Ŀ���=B.��Ŀ��� AND B.����ȼ�>=[1]  And Nvl(b.Ӧ�÷�ʽ,0)=1 " & _
            "ORDER BY A.�������"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(mrsParam("����ȼ�").Value))
    If rsTmp.BOF Then
        ShowSimpleMsg "�����µ�����������Ŀ�����ڻ�����Ŀ�����ã�"
        Exit Function
    End If

    Do While Not rsTmp.EOF

        Set cbrItem = mcbrMenuBar����.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendOther, zlCommFun.NVL(rsTmp("��¼��")), -1, False)
        cbrItem.Parameter = rsTmp.AbsolutePosition

        rsTmp.MoveNext
    Loop

    InitBodyLine = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'Private Function InitBody(ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal lng����id As Long) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim strSQL As String
'    Dim RS As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
'    Dim cbrItem As CommandBarControl
'    Dim intCount As Integer
'    Dim strDateFrom As String
'    Dim strDateTo As String
'    Dim strEnterDate As String
'    Dim intCol As Integer
'    Dim strCaption As String
'    Dim strParameter As String
'    Dim strNow As String
'    Dim strCut As String
'    Dim lngLoop As Long
'    Dim strTmp As String
'    Dim lnglast����id As Long
'
'    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
'    'ɾ������ҳ��˵���
'
'    mcbrToolBarҳ��.Delete
'    mcbrMenuBarҳ��.Delete
'
'    Set mcbrToolBarҳ�� = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "ҳ��", 5):  mcbrToolBarҳ��.BeginGroup = True
'    mcbrToolBarҳ��.IconId = conMenu_Edit_Modify
'    mcbrToolBarҳ��.Style = xtpButtonIconAndCaption
'
'    Set mcbrMenuBarҳ�� = mcbr�鿴.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewParent, "����ҳ��(&P)", 3)
'    mcbrMenuBarҳ��.BeginGroup = True
'
'    '
'    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Select ��Ժʱ��, ��Ժʱ��, 1 + Round((b.��Ժʱ�� - b.��Ժʱ��) / 7) As ҳ��" & vbNewLine & _
'                "  from (Select Min(��ʼʱ��) as ��Ժʱ��," & vbNewLine & _
'                "               Max(Nvl(��ֹʱ��, Sysdate)) as ��Ժʱ��" & vbNewLine & _
'                "          From ���˱䶯��¼" & vbNewLine & _
'                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]) b"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, lng��ҳid)
'    If rsTmp.BOF Then
'        MsgBox "�޲��˱���סԺ��¼��", vbExclamation, gstrSysName
'        Exit Function
'    End If
'
'    '
'    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Select 1 + Round((a.��ʼʱ�� - b.��Ժʱ��) / 7) As ��ʼҳ��,1 + Round((a.��ֹʱ�� - b.��Ժʱ��) / 7) As ����ҳ��,b.��Ժʱ��," & vbNewLine & _
'                "       ����id,c.����," & vbNewLine & _
'                "       ��ʼʱ��," & vbNewLine & _
'                "       ��ֹʱ��" & vbNewLine & _
'                "  from (Select ����id," & vbNewLine & _
'                "               Min(��ʼʱ��) as ��ʼʱ��," & vbNewLine & _
'                "               Max(Nvl(��ֹʱ��, Sysdate)) as ��ֹʱ��" & vbNewLine & _
'                "          From ���˱䶯��¼" & vbNewLine & _
'                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]" & vbNewLine & _
'                "         Group by ����id) a," & vbNewLine & _
'                "       (Select Min(��ʼʱ��) as ��Ժʱ��" & vbNewLine & _
'                "          From ���˱䶯��¼" & vbNewLine & _
'                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]) b,���ű� c Where c.ID=a.����id " & vbNewLine & _
'                " order by a.��ʼʱ��"
'    Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, lng��ҳid)
'
'    strEnterDate = Format(rsTmp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
'    For lngLoop = 0 To rsTmp("ҳ��").Value - 1
'
'        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
'        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
'        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
'            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
'        End If
'
'        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
'
'            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
'            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
'
'            RS.Filter = ""
'            RS.Filter = "��ʼҳ��<=" & lngLoop + 1 & " And ����ҳ��>=" & lngLoop + 1
'            If RS.RecordCount > 0 Then RS.MoveFirst
'            For intCol = 1 To RS.RecordCount
'
'                If strDateFrom < Format(RS("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
'                    strTmp = Format(RS("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
'                Else
'                    strTmp = strDateFrom
'                End If
'
'                If strDateTo > Format(RS("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
'                    strCaption = Format(RS("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss")
'                Else
'                    strCaption = strDateTo
'                End If
'
'                strCaption = Format(strTmp, "yyyy-MM-dd") & "��" & Format(strCaption, "yyyy-MM-dd")
'                strCaption = "��" & lngLoop + 1 & "ҳ��" & strCaption & "(" & RS("����").Value & ")"
'
'                Set cbrItem = mcbrMenuBarҳ��.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
'
'                '��Ժʱ��;����id;��ʼʱ��;����ʱ��;
'                cbrItem.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
'
'                Set cbrItem = mcbrToolBarҳ��.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
'                cbrItem.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
'
'                lnglast����id = RS("����ID").Value
'
'                RS.MoveNext
'
'                strParameter = cbrItem.Parameter
'            Next
'        End If
'
'    Next
'
'    If strParameter <> "" Then Call BodyEdit.zlMenuClick("װ������", strParameter)
'
'    InitBody = True
'End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
       
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "��������(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�ָ�����(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    Set mcbrMenuBar�༭ = cbrMenuBar
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "�趨��ʼ����(&B)")
        
        Set mcbrMenuBar���� = .Add(xtpControlPopup, conMenu_Edit_Modify, "��������(&D)")
        mcbrMenuBar����.BeginGroup = True
        mcbrMenuBar����.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_SendOther, "��", -1, False
        
        Set cbrPop = .Add(xtpControlPopup, conMenu_Edit_Append, "���⴦��(&S)")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "ʧ����ٸ�(&1)", -1, False): cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "�೦(&2)", -1, False):  cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "�೦����й(&3)", -1, False):  cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "����(&4)", -1, False):   cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "��������(&5)", -1, False):   cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "�����Ŀ(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ɾ����Ŀ(&R)")
        
        '
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "���ü�¼(&E)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "�����¼(&U)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��������/����(&W)"): cbrControl.IconId = 1
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "�������/����(&C)"): cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���Ժϸ�(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ȡ������(&B)")
        
        Set cbrPop = .Add(xtpControlPopup, conMenu_View_ToolBar, "�Զ���ȡ(&A)"): cbrPop.BeginGroup = True: cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Price, "��ȡ����(&1)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Price, "��ȡ����/����(&2)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
        
    End With

    Set mcbr�鿴 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With mcbr�鿴.CommandBar.Controls
                
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
                
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
'        Set mcbrMenuBarҳ�� = .Add(xtpControlPopup, conMenu_Edit_NewParent, "����ҳ��(&P)")
'        mcbrMenuBarҳ��.BeginGroup = True
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
   
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(zlDatabase.GetPara("���ҷ���", glngSys, 1255, "��  ��"))
    If mstrFindKey = "" Then mstrFindKey = "��  ��"
        
    Set mobjFindKey = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.ToolTipText = "��ݼ�:F4"
    mobjFindKey.Style = xtpButtonIconAndCaption
    mobjFindKey.flags = xtpFlagRightAlign
    
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.��  ��"): cbrControl.Parameter = "��  ��"
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.סԺ��"): cbrControl.Parameter = "סԺ��"
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.���￨"): cbrControl.Parameter = "���￨"

    Set cbrCustom = cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picCustom.hWnd
    txt.ToolTipText = "���Ҳ���(F3)"
    cbrCustom.flags = xtpFlagRightAlign

    Set cbrControl = cbsThis.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ����")
    cbrControl.ToolTipText = "ǰһ����(Ctrl+Left)"
    cbrControl.flags = xtpFlagRightAlign
    cbrControl.Style = xtpButtonIcon

    Set cbrControl = cbsThis.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ����")
    cbrControl.ToolTipText = "��һ����(Ctrl+Right)"
    cbrControl.flags = xtpFlagRightAlign
    cbrControl.Style = xtpButtonIcon
    
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�ָ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    
    '��λ������
    '------------------------------------------------------------------------------------------------------------------
    
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FALT, Asc("1"), conMenu_Edit_Append * 10 + 1
        .Add FALT, Asc("2"), conMenu_Edit_Append * 10 + 2
        .Add FALT, Asc("3"), conMenu_Edit_Append * 10 + 3
        .Add FALT, Asc("4"), conMenu_Edit_Append * 10 + 4
        .Add FALT, Asc("5"), conMenu_Edit_Append * 10 + 5
        .Add 0, VK_DELETE, conMenu_Edit_Untread
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add 0, VK_F1, conMenu_Help_Help
                
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
        
    End With
    
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub BodyEdit_PromptInfo(ByVal strInfo As String)
    stbThis.Panels(2).Text = strInfo
End Sub

Private Sub BodyEdit_RButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub
    
    '��װ�Ҽ��˵�
    If mcbrMenuBar��λ Is Nothing Then Exit Sub
    If mcbrMenuBar��λ.CommandBar.Controls.Count = 0 Then Exit Sub
    Set cbrMenuBar = mcbrMenuBar��λ
    Set cbrPopupBar = cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.Id, cbrControl.Caption)
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
    
End Sub

Private Sub BodyEdit_SelectScale(ByVal intScale As Integer)
    Call AddActiveMenu
End Sub

'######################################################################################################################
'�ؼ��¼�

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_Tool_Option
            
            If Control.Parameter = "" Then
                Control.Parameter = "1"
            Else
                Control.Parameter = ""
            End If
            
        Case conMenu_File_PrintSet
            
            On Error Resume Next
            frmPrintSet.mbytMode = 1
            frmPrintSet.mstrPrivs = mstrPrivs
            frmPrintSet.Show 1, Me
            
        Case conMenu_File_Preview
            
            Call PrintData(2)
            
        Case conMenu_File_Print
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_Notify    '�趨���µ���ʼ����(�Ѵ����������ݵĲ������趨)
            Dim strParam As String
            
            If Not BodyEdit.zlMenuClick("�趨��ʼ����") Then Exit Sub
            '����������ID,��ҳID,����ID,����ID,��Ժ��־;�༭��־;Ӥ��
            strParam = Val(mrsParam("����id")) & ";" & Val(mrsParam("��ҳid")) & ";" & Val(mrsParam("����id")) & ";" & Val(mrsParam("��Ժ")) & ";" & Val(mrsParam("�༭").Value) & ";" & Val(mrsParam("Ӥ��").Value)
            Call BodyEdit.zlMenuClick("��ʼ����", strParam)
        
        Case conMenu_Edit_Adjust
            
            If BodyEdit.CurPostion >= 0 Then Call BodyEdit.zlMenuClick("��д��¼��")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            
            If BodyEdit.CurPostion >= 0 Then Call BodyEdit.zlMenuClick("�����¼��")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
    
            Call BodyEdit.zlMenuClick("��д������")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
        
            Call BodyEdit.zlMenuClick("���������")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Append            '�����Ŀ
            Call BodyEdit.zlMenuClick("�����Ŀ")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Stop              'ɾ����Ŀ
            Call BodyEdit.zlMenuClick("ɾ����Ŀ")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 1, conMenu_Edit_Compend * 10 + 2, conMenu_Edit_Compend * 10 + 3
            
            mstr���²�λ = Control.Parameter
            
            BodyEdit.���²�λ = mstr���²�λ
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 5, conMenu_Edit_Compend * 10 + 6
            
            mstr������ʽ = Control.Parameter
            
            BodyEdit.������ʽ = mstr������ʽ
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 8
            
            mstr���� = IIf(Control.Checked = False, "����", "")
            
            BodyEdit.������ʽ = mstr����
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Reuse
            If BodyEdit.zlMenuClick("�ָ�����") Then
    
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit
            
            Call BodyEdit.zlMenuClick("���Ժϸ�")
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Blankoff
            
            Call BodyEdit.zlMenuClick("ȡ������")
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location
            
            Call LocationObj(txt)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward
            
            If mlngRowNum = 1 Then mlngRowNum = mrsPatient.RecordCount + 1
            
            mrsPatient.Filter = ""
            mrsPatient.Filter = "ID<" & mlngRowNum
            If mrsPatient.RecordCount > 0 Then
                mrsPatient.MoveLast
                mlngRowNum = Val(mrsPatient("ID").Value)
                txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
                mrsParam("����id").Value = Val(mrsPatient("����id").Value)
                mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
                mrsParam("Ӥ��").Value = 0
                Select Case CStr(mrsPatient("����").Value)
                Case "����", "��������", "��Ժ����"
                    mrsParam("��Ժ").Value = 1
                Case Else
                    mrsParam("��Ժ").Value = 0
                End Select
                
                Call OpenPatientMap
                txt.Tag = ""
            End If
            mrsPatient.Filter = ""
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward
            
            If mlngRowNum = mrsPatient.RecordCount Then mlngRowNum = 0
            
            mrsPatient.Filter = ""
            mrsPatient.Filter = "ID>" & mlngRowNum
            If mrsPatient.RecordCount > 0 Then
                mrsPatient.MoveFirst
                mlngRowNum = Val(mrsPatient("ID").Value)
                txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
                mrsParam("����id").Value = Val(mrsPatient("����id").Value)
                mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
                mrsParam("Ӥ��").Value = 0
                Select Case CStr(mrsPatient("����").Value)
                Case "����", "��������", "��Ժ����"
                    mrsParam("��Ժ").Value = 1
                Case Else
                    mrsParam("��Ժ").Value = 0
                End Select
                
                Call OpenPatientMap
                txt.Tag = ""
            End If
            mrsPatient.Filter = ""
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Save
            '��������
            
            If BodyEdit.zlMenuClick("��������") Then
                mblnChanged = True
            End If
            
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Price
            
            '��������
            Select Case Control.Parameter
            Case "����"
                mblnChanged = BodyEdit.zlMenuClick("��������")
            Case "����"
                mblnChanged = BodyEdit.zlMenuClick("��ȡ������")
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Append * 10 + 1

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("�ٸ�")

        Case conMenu_Edit_Append * 10 + 2

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("�೦")

        Case conMenu_Edit_Append * 10 + 3

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("�೦����й")

        Case conMenu_Edit_Append * 10 + 4

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("����")
            
        Case conMenu_Edit_Append * 10 + 5

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("��������")
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Jump
            
            Call BodyEdit.zlMenuClick("װ������", Control.Parameter)
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SendOther
            
            Call BodyEdit.zlMenuClick("��������", Val(Control.Parameter))
            
            Call AddActiveMenu
            
            cbsThis.RecalcLayout
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_Save, conMenu_Edit_Reuse, conMenu_Edit_Price
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit                 '���Ժϸ�
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.AllowAudit)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Blankoff              'ȡ������
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.AllowUnAudit)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
    
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And Val(BodyEdit.GetUpObj.ColData(BodyEdit.GetUpObj.Col)) > 0)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append            '�����Ŀ
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Stop              'ɾ����Ŀ
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 1, conMenu_Edit_Compend * 10 + 2, conMenu_Edit_Compend * 10 + 3    '����/Ҹ��/����
    
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1) And BodyEdit.������Ŀ
        Control.Checked = (Control.Parameter = mstr���²�λ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 5, conMenu_Edit_Compend * 10 + 6                                   '��������/����������
    
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1) And BodyEdit.������Ŀ
        Control.Checked = (Control.Parameter = mstr������ʽ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 8                                                                  '����ʹ������
    
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1) And BodyEdit.������Ŀ
        Control.Checked = (mstr���� = "����")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Adjust, conMenu_Edit_Untread
        
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.CurPostion >= 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 1

'        Control.Checked = (BodyEdit.mbytSpecChar = 1)
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.�Ƿ�����Ŀ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 2

'        Control.Checked = (BodyEdit.mbytSpecChar = 2)
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.�Ƿ�����Ŀ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 3

'        Control.Checked = (BodyEdit.mbytSpecChar = 3)
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.�Ƿ�����Ŀ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 4
    
'        Control.Checked = (BodyEdit.mbytSpecChar = 4)
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.�Ƿ��Һ��Ŀ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 5
    
'        Control.Checked = (BodyEdit.mbytSpecChar = 5)
        Control.Enabled = (Val(mrsParam("�༭").Value) = 1 And BodyEdit.�Ƿ��Һ��Ŀ)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If Control.Parameter = "" Then
            Control.Checked = True
        Else
            Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = BodyEdit.Page)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SendOther
        
        Control.Checked = (Val(Control.Parameter) = BodyEdit.LineType)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        Control.Checked = (mstrFindKey = Control.Parameter)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
'        Control.Enabled = (mlngRowNum > 1)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward, conMenu_View_Backward
        
        Control.Enabled = (mrsPatient.RecordCount > 1)
        
    End Select
End Sub

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset

    '------------------------------------------------------------------------------------------------------------------
    mrsPatient.Filter = ""
    If mrsPatient.RecordCount > 0 Then
        mrsPatient.MoveFirst
        If ShowTxtSelDialog(Me, txt, "����,1200,0,0;����,1200,0,1;�Ա�,600,0,0;����,1800,0,0;סԺ��,1080,0,0", Me.Name & "\�����嵥ѡ��", "�������ѡ��һ�����ˡ�", mrsPatient, rs, 5600, 4500, , CStr(mlngRowNum), 2, True) Then
            
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(rs("����").Value)
            
            
            mrsParam("����id").Value = Val(rs("����id").Value)
            mrsParam("��ҳid").Value = Val(rs("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            Select Case CStr(rs("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select
            
            Call OpenPatientMap
            
            txt.Tag = ""
        End If
    End If
    mrsPatient.Filter = ""
    
    Call LocationObj(txt)

    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub Form_Load()
        
    Call InitCommonControls
    
    If mblnChildForm Then
'        Call RestoreWinState(Me, App.ProductName, "ChildForm")
    Else
        Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Call zlDatabase.SetPara("���ҷ���", mstrFindKey, glngSys, 1255)
    
    UnHook Me.hWnd
    
    If mblnChildForm Then
'        Call SaveWinState(Me, App.ProductName, "ChildForm")
    Else
        Call SaveWinState(Me, App.ProductName)
    End If
    
    
    
    Set mrsPatient = Nothing
    Set mobjFindKey = Nothing
    mblnShowing = False
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub BodyEdit_DbClickCur()
    
    Call BodyEdit.zlMenuClick("��д��¼��")
        
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim bytMode As Byte
    Dim lng����ID As Long
    Dim strInput As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" And txt.Text <> "" Then
            If InStr(txt.Text, "'") Then
                ShowSimpleMsg "������������зǷ��ַ� ' ��"
                Exit Sub
            End If
            
            Select Case mstrFindKey
'            Case "����id"
'                strInput = "����id=" & Val(txt.Text)
'                bytMode = 2
'            Case "�����"
'                strInput = "�����=" & Val(txt.Text)
'                bytMode = 4
            Case "��  ��"
                strInput = "����='" & Trim(txt.Text) & "'"
                bytMode = 5
            Case "סԺ��"
                strInput = "סԺ��=" & Val(txt.Text)
                bytMode = 3
            Case "���￨"
                strInput = "���￨��='" & Trim(txt.Text) & "'"
                bytMode = 1
            End Select
                        
        End If

    ElseIf mstrFindKey = "���￨" And txt.Tag = "Changed" And txt.Text <> "" Then
        If Len(txt.Text) = mint���￨���볤�� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt.Text <> "" Then
            If KeyAscii <> 13 Then
                txt.Text = txt.Text & Chr(KeyAscii)
                txt.SelStart = Len(txt.Text)
                KeyAscii = 0
            End If

            strInput = "���￨��='" & Trim(txt.Text) & "'"
            bytMode = 1
        End If
    End If
    
    If strInput <> "" Then
        txt.Tag = ""
        mrsPatient.Filter = ""
        mrsPatient.Filter = strInput
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            lng����ID = Val(mrsPatient("����id").Value)
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
            txt.Tag = ""
            
            mrsParam("����id").Value = Val(mrsPatient("����id").Value)
            mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            Select Case CStr(mrsPatient("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select
            
            Call OpenPatientMap
        Else
            ShowSimpleMsg "û���ҵ����������Ĳ��ˣ�"
            txt.Text = mstrSvr����
        End If
        mrsPatient.Filter = ""

        Call LocationObj(txt)
        
    End If

    Exit Sub

errHand:
End Sub

Private Sub AddActiveMenu()
    '------------------------------------------------------------
    '������Ŀ��Ӳ˵�(������������������²�λ;����Ǻ��������Ӻ�����ʽ)
    Dim varTmp As Variant
    Dim rs As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    If Not mcbrMenuBar��λ Is Nothing Then
        If mcbrMenuBar��λ.CommandBar.Controls.Count <> 0 Then
            Call mcbrMenuBar��λ.CommandBar.Controls.DeleteAll
            Call mcbrMenuBar�༭.CommandBar.Controls.Item(2).Delete
        End If
    End If
    If BodyEdit.������Ŀ Then
        Set mcbrMenuBar��λ = mcbrMenuBar�༭.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "���²�λ(&T)", 2)
        gstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1)
        If rs.BOF = False Then
            varTmp = Split(zlCommFun.NVL(rs("��¼��").Value, "��,��,��"), ",")
        Else
            varTmp = Split("��,��,��", ",")
        End If
        
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 1, "����" & varTmp(0) & "(&1)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 2, "Ҹ��" & varTmp(1) & "(&2)", -1, False): cbrControl.Parameter = "Ҹ��": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 3, "����" & varTmp(2) & "(&3)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
    End If
    
    If BodyEdit.������Ŀ Then
        Set mcbrMenuBar��λ = mcbrMenuBar�༭.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "������ʽ(&T)", 2)
        gstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 3)
        If rs.BOF = False Then
            varTmp = zlCommFun.NVL(rs("��¼��").Value, "��")
        Else
            varTmp = "��"
        End If
        
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 5, "��������" & varTmp & "(&1)", -1, False): cbrControl.Parameter = "��������": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 6, "������ (&2)", -1, False): cbrControl.Parameter = "������": cbrControl.IconId = 1
    End If

    If BodyEdit.������Ŀ Then
        Set mcbrMenuBar��λ = mcbrMenuBar�༭.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "������ʽ(&T)", 2)
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 8, "ʹ������" & "(&1)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
    End If
End Sub
