VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_SpeExam 
   Caption         =   "�ؼ�����"
   ClientHeight    =   7980
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10260
   Icon            =   "frmPatholRequisition_SpeExam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10260
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7440
      Width           =   2145
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "�� ��(&A)"
      Height          =   400
      Left            =   7320
      TabIndex        =   22
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox chkPriceState 
      Caption         =   "�貹��"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   7560
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�� ��(&E)"
      Height          =   400
      Left            =   8640
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame framAntibody 
      Caption         =   "�ؼ���Ŀ"
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   9975
      Begin VB.PictureBox picMealClass 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   7080
         ScaleHeight     =   5295
         ScaleWidth      =   2655
         TabIndex        =   7
         Top             =   360
         Width           =   2655
         Begin VB.ListBox lstMealClass 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4050
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   2655
         End
         Begin VB.ComboBox cbxMealClass 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label labMealClass 
            Caption         =   "�ײ����"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txtAntibodyName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   0
         ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ��"
         Top             =   5445
         Width           =   2625
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8705
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ƹ��ˣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ��"
         Top             =   5520
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cbxMaterial 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   3225
      End
      Begin VB.ComboBox cbxSpeExamType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholRequisition_SpeExam.frx":179A
         Left            =   5520
         List            =   "frmPatholRequisition_SpeExam.frx":179C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   3225
      End
      Begin VB.ComboBox cbxSpeExamDetails 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholRequisition_SpeExam.frx":179E
         Left            =   1080
         List            =   "frmPatholRequisition_SpeExam.frx":17A0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   3225
      End
      Begin VB.Label labMaterial 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ŀ��ţ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label labSpeExamType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ؼ����ͣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   20
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   200
      End
      Begin VB.Label Label5 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   240
         Width           =   200
      End
      Begin VB.Label labDescription 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   17
         Top             =   810
         Width           =   900
      End
      Begin VB.Label labSpeexamDetails 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ؼ�ϸĿ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   900
      End
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   7440
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   74317827
      CurrentDate     =   40646.4399652778
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   8640
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labRequestDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽʦ��"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3000
      TabIndex        =   26
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʱ�䣺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   7500
      Width           =   540
   End
End
Attribute VB_Name = "frmPatholRequisition_SpeExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentRequest As ucFlexGrid
Private mufgParentContext As ucFlexGrid

Private mblnIsBuZuo As Boolean
Private mfrmOwner As Form

Private mrsMealLink As New ADODB.Recordset

Private mlngCurRequestId As Long
Private mlngPatholAdviceId As Long


Public blnIsOk As Boolean


Public Function ShowSpeExamRequestWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, ByVal lngRequestId As Long, owner As Form, Optional ByVal blnIsBuZuo As Boolean = False) As Boolean
'��ʾ�ؼ����봰��
    Set mufgParentRequest = ufgParentRequestGrid
    Set mufgParentContext = ufgParentContextGrid

    Set mfrmOwner = owner
    
    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurRequestId = lngRequestId
    mblnIsBuZuo = blnIsBuZuo
    blnIsOk = False
        
    '����Ŀ���Ϣ
    Call LoadMaterialInf
    
    dtpRequestTime.value = zlDatabase.Currentdate
    txtRequestDoctor.Text = UserInfo.����
    
    If lngRequestId > 0 Then
        Select Case ufgParentRequestGrid.Text(ufgParentRequestGrid.SelectionRow, gstrRequisition_��������)
            Case "�����黯"
                cbxSpeExamType.ListIndex = 0
            Case "����Ⱦɫ"
                cbxSpeExamType.ListIndex = 1
            Case "���Ӳ���"
                cbxSpeExamType.ListIndex = 2
        End Select
        
        Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
        
        Select Case ufgParentRequestGrid.Text(ufgParentRequestGrid.SelectionRow, gstrRequisition_����ϸĿ)
            Case "����", "ӫ��"
                cbxSpeexamDetails.ListIndex = 0
            Case "��ҩ��ҩ", "��ͨ"
                cbxSpeexamDetails.ListIndex = 1
        End Select
        
    End If
    
    '�����ؼ��������
    Call AdjustSpeExamFace(lngRequestId <= 0)
    
    Call Me.Show(1, owner)
End Function


Private Sub AdjustSpeExamFace(ByVal blnIsNewRequest As Boolean)
'�����ؼ����
    labRequestTime.Enabled = blnIsNewRequest
    dtpRequestTime.Enabled = blnIsNewRequest
    
    labSpeExamType.Enabled = blnIsNewRequest
    cbxSpeExamType.Enabled = blnIsNewRequest
    
    labSpeexamDetails.Enabled = blnIsNewRequest
    cbxSpeexamDetails.Enabled = blnIsNewRequest
    
    labRequestDoctor.Enabled = blnIsNewRequest
    txtRequestDoctor.Enabled = blnIsNewRequest
    
    labDescription.Enabled = blnIsNewRequest
    txtDescription.Enabled = blnIsNewRequest
    
    cbxSpeExamType.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
    cbxSpeexamDetails.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
    txtDescription.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
End Sub


Private Sub AdjustFace()
    framAntibody.Height = Me.Height - framAntibody.Top - cmdExit.Height - 800
    framAntibody.Width = Me.Width - (framAntibody.Left * 2) - 120
    
    ufgData.Left = 120
    ufgData.Top = 240
    
    ufgData.Height = framAntibody.Height - txtAntibodyName.Height - 480
    ufgData.Width = framAntibody.Width - picMealClass.Width - 360
    
    picMealClass.Left = ufgData.Left + ufgData.Width + 120
    picMealClass.Top = ufgData.Top
    picMealClass.Height = ufgData.Height + txtAntibodyName.Height + 120
    
    lstMealClass.Height = picMealClass.Height - lstMealClass.Top - txtAntibodyName.Height
        
    
    labFilter.Left = 120
    
    
    txtAntibodyName.Left = labFilter.Left + labFilter.Width + 60
    txtAntibodyName.Top = ufgData.Top + ufgData.Height + 120
    txtAntibodyName.Width = ufgData.Width - txtAntibodyName.Left + 120
    
    labFilter.Top = txtAntibodyName.Top + 60
    
    
    cmdExit.Left = framAntibody.Width - cmdExit.Width + framAntibody.Left
    cmdExit.Top = framAntibody.Top + framAntibody.Height + 120
    
    cmdApply.Left = cmdExit.Left - cmdApply.Width - 120
    cmdApply.Top = cmdExit.Top
    
    
    chkPriceState.Left = cmdApply.Left - chkPriceState.Width - 120
    chkPriceState.Top = cmdExit.Top + 90
    
    labRequestTime.Left = 120
    labRequestTime.Top = chkPriceState.Top
    
    dtpRequestTime.Left = labRequestTime.Left + labRequestTime.Width + 60
    dtpRequestTime.Top = cmdExit.Top + 30
    
    labRequestDoctor.Left = dtpRequestTime.Left + dtpRequestTime.Width + 240
    labRequestDoctor.Top = labRequestTime.Top
    
    txtRequestDoctor.Left = labRequestDoctor.Left + labRequestDoctor.Width + 60
    txtRequestDoctor.Top = cmdExit.Top + 30
End Sub



Private Sub LoadMaterialInf()
'����Ŀ���Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    
    strSQL = "select �Ŀ�ID,���,�걾���� from ����ȡ����Ϣ where ����ҽ��ID=[1] and ȷ��״̬=1"
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatholAdviceId)
    
    Call cbxMaterial.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call cbxMaterial.AddItem(Nvl(rsData!�Ŀ�id) & " (�Ŀ�ţ�" & rsData!��� & " �걾����" & rsData!�걾���� & ")")

        Call rsData.MoveNext
    Loop
    
End Sub



Private Sub LoadSpeExamType()
'���뿹������
    cbxSpeExamType.Clear
    
    Call cbxSpeExamType.AddItem("0-�����黯")
    Call cbxSpeExamType.AddItem("1-����Ⱦɫ")
    Call cbxSpeExamType.AddItem("2-���Ӳ���")
    
    cbxSpeExamType.ListIndex = 0
End Sub



Private Sub LoadSpeExamDetails(ByVal lngSpeExamType As Long)
'�����ؼ���ϸ
    cbxSpeexamDetails.Clear
    
'    Call cbxSpeExamDetails.AddItem("")
    
    If lngSpeExamType = TSpeexamType.stMianyi Then
        Call cbxSpeexamDetails.AddItem("1-����(����)")
        Call cbxSpeexamDetails.AddItem("2-����(��ҩ��ҩ)")
        
        cbxSpeexamDetails.ListIndex = 1
    ElseIf lngSpeExamType = TSpeexamType.stFenzi Then
        Call cbxSpeexamDetails.AddItem("1-����(ӫ��)")  '��Ӧ 3
        Call cbxSpeexamDetails.AddItem("2-����(��ͨ)")  '��Ӧ 4
        
        cbxSpeexamDetails.ListIndex = 0
    End If
End Sub



Private Sub LoadMealLinkData()
'�����ײ͹�������
    Dim strSQL As String
    
    '��ȡ�ײ͹�������
    strSQL = "select a.�ײ�ID, a.�ײ�����, b.����id, b.����˳�� from �����ײ���Ϣ a, �����ײ͹��� b where a.�ײ�id=b.�ײ�id"
    Set mrsMealLink = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
End Sub


Private Sub LoadAntibodyMeal(ByVal strMealClass As String)
'���뿹���ײ�
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    '��ȡ�ײ�����
    strSQL = "select �ײ����� from �����ײ���Ϣ " & IIf(strMealClass <> "", " where �ײ����=[1]", "") & " order by �ײ�����"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMealClass)
    
    Call lstMealClass.Clear
    
    Call lstMealClass.AddItem("")
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call lstMealClass.AddItem(Nvl(rsData!�ײ�����))
        rsData.MoveNext
    Loop
    
End Sub


Private Sub InitAntibodyList()
'��ʼ��������Ϣ�б�
    Dim strTemp As String
    

    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�ؼ쿹���б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrRequestAntibodyCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
        '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrRequestAntibodyCols
    ufgData.ColConvertFormat = gstrRequestAntibodyConvertFormat
End Sub



Private Sub QueryAntibodyData()
'��ѯ��������
    Dim strSQL As String
    
    strSQL = "select a.����id, a.��������,a.ʹ���˷�,a.�����˷�,a.��������,a.��Ч��,a.��������, '' as ��Ŀ˳�� " & _
                " from ��������Ϣ a where a.ʹ��״̬=1 order by a.�������� "
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub GetMealAntibodyIds(ByVal strMealName As String, ByRef strAntibodyIds As String, ByRef strAntibodyOrder As String)
'ȡ���ײ������Ŀ���Id��
    
    strAntibodyIds = ""
    strAntibodyOrder = ""
    
    
    mrsMealLink.Filter = "�ײ�����='" & strMealName & "'"
    If mrsMealLink.RecordCount <= 0 Then Exit Sub
    
    Do While Not mrsMealLink.EOF
        If strAntibodyIds <> "" Then strAntibodyIds = strAntibodyIds & " or "
        strAntibodyIds = strAntibodyIds & "����Id=" & mrsMealLink!����ID
        strAntibodyOrder = strAntibodyOrder & mrsMealLink!����ID & ":" & _
            String(5 - Len("" & mrsMealLink!�ײ�ID & ""), "0") & mrsMealLink!�ײ�ID & String(5 - Len("" & mrsMealLink!����˳�� & ""), "0") & mrsMealLink!����˳�� & ";"
        
        Call mrsMealLink.MoveNext
    Loop

End Sub


Private Sub LoadAntibodyDataToFace(ByVal strAntibodyMeal As String)
'��ȡ������Ϣ��������ʾ
    Dim strCurMealAntibodyIds As String
    Dim strCurAntibodyOrders As String
    Dim strCurAntibodyOrder As String
    Dim i As Long
    
    'ȡ�õ�ǰ�ײ��������ײ�ID
    If strAntibodyMeal <> "" Then
        Call GetMealAntibodyIds(strAntibodyMeal, strCurMealAntibodyIds, strCurAntibodyOrders)
        ufgData.AdoFilter = IIf(strCurMealAntibodyIds <> "", strCurMealAntibodyIds, "����ID=-1")
    Else
        ufgData.AdoFilter = ""
    End If
    
    Call ufgData.RefreshData
    
    'д�뵱ǰ�ײ͵Ŀ���˳��
    If strCurAntibodyOrders = "" Then Exit Sub
    
    On Error Resume Next
    For i = 1 To ufgData.GridRows - 1
        strCurAntibodyOrder = ufgData.KeyValue(i)
        strCurAntibodyOrder = Mid(strCurAntibodyOrders, InStr(strCurAntibodyOrders, strCurAntibodyOrder & ":") + Len(strCurAntibodyOrder) + 1, 100)
        strCurAntibodyOrder = Mid(strCurAntibodyOrder, 1, InStr(strCurAntibodyOrder, ";") - 1)
        
        '�����ײ��µĿ���˳��
        ufgData.Text(i, gstrRequestAntibody_��Ŀ˳��) = strCurAntibodyOrder
    Next i
    
    '���ײ͵Ŀ���˳������
    Call ufgData.Sort(ufgData.GetColIndex(gstrRequestAntibody_��Ŀ˳��))
End Sub



Private Function GetSpecimenAntibodyIds(ByVal strMaterialId As String)
'��ȡ�걾��Ӧ�Ŀ���Id
    Dim i As Long
    Dim strIds As String
    
    'strIds ������ʽΪ",asf,aat,aft,bbe,"
    
    strIds = ""
    For i = 1 To mufgParentContext.GridRows - 1
        If Not mufgParentContext.IsEmptyKey(i) Then
            If Val(mufgParentContext.Text(i, gstrRequest_SpeExam_�Ŀ��)) = Val(strMaterialId) Then
                strIds = strIds & "," & mufgParentContext.Text(i, gstrRequest_SpeExam_��������)
            End If
        End If
    Next i
    
    If strIds <> "" Then strIds = strIds & ","
    
    GetSpecimenAntibodyIds = strIds
End Function


Private Function HideSpecimenAntibody(ByVal strSpecimenAntibodyIds As String)
'���ر걾�Ѿ�ѡ��Ŀ���
    Dim i As Long
    
    For i = 1 To ufgData.GridRows - 1
        If UCase(strSpecimenAntibodyIds) Like "*," & UCase(ufgData.Text(i, gstrRequestAntibody_��������)) & ",*" Then
            ufgData.RowHidden(i) = True
        End If
    Next i
End Function


Private Sub cbxMaterial_Click()
On Error GoTo ErrHandle
    Dim strSpecimenAntibodyIds As String
    
    '�ָ��б������ص�����
    Call ufgData.RestoreList


    
'    If mblnIsBuZuo Then
        strSpecimenAntibodyIds = GetSpecimenAntibodyIds(GetSelectMaterialNum)
    
        '���ض�Ӧ�Ŀ���
        Call HideSpecimenAntibody(strSpecimenAntibodyIds)
'    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxMealClass_Click()
On Error GoTo ErrHandle

    Call LoadAntibodyMeal(cbxMealClass.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxSpeExamType_Click()
On Error GoTo ErrHandle
    '�����ؼ�ϸĿ
    Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
    
    Exit Sub
ErrHandle:
    If ErrCenter() = True Then Resume
End Sub

Private Sub cmdApply_Click()
'�����ؼ�����
On Error GoTo ErrHandle
    '���������Ч����ֱ���˳�������Ҫ�ٽ�����ʾ
    If Not CheckDataIsValid Then Exit Sub
    
    '�ж�������ϸ�б��Ƿ�Ϊ�ؼ���Ŀ��ϸ�б�
    If mufgParentContext.GetColIndex(gstrRequest_SpeExam_��������) < 0 Then
        mufgParentContext.ColNames = gstrRequest_SpeExam_Cols
        mufgParentContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
        
        
        '�л�������Ŀ��ƽ���
        Call mfrmOwner.ChangeControlFace(0)
    End If
    
    '�����ؼ�����
    Call SaveSpeExamRequest
    
    blnIsOk = True
    
    If MsgBoxD(Me, "��������ɣ��Ƿ������ӣ�", vbYesNo, Me.Caption) = vbNo Then
        Call Me.Hide
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
'�˳�����
On Error Resume Next
    blnIsOk = False
    Call Me.Hide
End Sub



Private Function CheckDataIsValid() As Boolean
'���¼�������Ƿ���Ч
    CheckDataIsValid = False
    
    '�ж��Ƿ�ѡ���˲Ŀ�
    If Trim(cbxMaterial.Text) = "" Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����ؼ촦��ĲĿ顣", vbInformation, Me.Caption)
        cbxMaterial.SetFocus
        
        Exit Function
    End If
    
    
    If mlngCurRequestId <= 0 Then
        '�ж��Ƿ�ѡ�����ؼ�����  (ֻ�����������Ҫ�����ж�)
        If Trim(cbxSpeExamType.Text) = "" Then
            Call MsgBoxD(Me, "��ѡ���ؼ�Ĵ������͡�", vbInformation, Me.Caption)
            cbxSpeExamType.SetFocus
            
            Exit Function
        End If
    End If
    
    '�ж��Ƿ�ѡ��������
    If Not ufgData.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫʹ�õĿ������ݡ�", vbInformation, Me.Caption)
        ufgData.SetFocus
        
        Exit Function
    End If
    
    CheckDataIsValid = True

End Function



Private Sub SaveSpeExamRequest()
'�����ؼ�����
    Dim lngNewRow As Long
    
    Dim i As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnContinueAdd As Boolean
    Dim lngSpeexamDetails As Long
    Dim blnIsInvalidDate As Boolean
    Dim blnIsInvalidCount As Boolean
    
    lngSpeexamDetails = 0
    
    '��ȡ��ǰ�ؼ�ϸĿ
    Select Case Val(cbxSpeExamType.Text)
        Case 0
            lngSpeexamDetails = Val(cbxSpeexamDetails.Text)
        Case 1
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxSpeexamDetails.Text) > 0, Val(cbxSpeexamDetails.Text) + 2, 0)
    End Select
        
        
        
    If mlngCurRequestId <= 0 Then
        
        '��Ӽ��������Ϣ
        strSQL = "select Zl_��������_����([1],[2],[3],[4],[5],[6],[7]) as ����ֵ from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                mlngPatholAdviceId, _
                                                txtRequestDoctor.Text, _
                                                CDate(dtpRequestTime.value), _
                                                Val(cbxSpeExamType.Text), _
                                                IIf(chkPriceState.value <> 0, 1, 0), _
                                                lngSpeexamDetails, _
                                                txtDescription.Text)
                                                
        If rsData.RecordCount <= 0 Then
            Call err.Raise(0, "SaveSpeExamRequest", "δ�ɹ���ȡ�����������ID,����ʧ�ܡ�")
            Exit Sub
        End If

        '���ý�����Ϣ
        lngNewRow = mufgParentRequest.NewRow
        
        mufgParentRequest.Text(lngNewRow, gstrRequisition_����ID) = rsData!����ֵ
        mufgParentRequest.Text(lngNewRow, gstrRequisition_������) = txtRequestDoctor.Text
        mufgParentRequest.Text(lngNewRow, gstrRequisition_��������) = Trim(Substr(cbxSpeExamType.Text, InStr(1, cbxSpeExamType.Text, "-") + 1, 10))
        mufgParentRequest.Text(lngNewRow, gstrRequisition_����ϸĿ) = Decode(lngSpeexamDetails, 1, "����", 2, "��ҩ��ҩ", 3, "ӫ��", 4, "��ͨ", "��")
        mufgParentRequest.Text(lngNewRow, gstrRequisition_����״̬) = IIf(chkPriceState.value <> 0, "�貹��", "��")
        mufgParentRequest.Text(lngNewRow, gstrRequisition_����ʱ��) = dtpRequestTime.value
        mufgParentRequest.Text(lngNewRow, gstrRequisition_����ʱ��) = dtpRequestTime.value
        mufgParentRequest.Text(lngNewRow, gstrRequisition_��������) = txtDescription.Text
        mufgParentRequest.Text(lngNewRow, gstrRequisition_��ǰ״̬) = "������"
                                                
        mlngCurRequestId = Val(Nvl(rsData!����ֵ))
        
        '��λ��������
        Call mufgParentRequest.LocateRow(lngNewRow)
        
        '���ԭ��������Ŀ����
        Call mufgParentContext.ClearListData
    End If
    
    
    
    '����ؼ�������Ŀ
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) And Not ufgData.RowHidden(i) Then
        
            blnContinueAdd = True
            
            blnIsInvalidCount = Val(ufgData.Text(i, gstrRequestAntibody_ʹ���˷�)) <= Val(ufgData.Text(i, gstrRequestAntibody_�����˷�))
            blnIsInvalidDate = zlDatabase.Currentdate > CDate(IIf(ufgData.Text(i, gstrRequestAntibody_��������) = "", "3000-01-01", ufgData.Text(i, gstrRequestAntibody_��������))) _
                    Or zlDatabase.Currentdate > DateAdd("m", _
                        Val(IIf(ufgData.Text(i, gstrRequestAntibody_��Ч��) = "", 2400, ufgData.Text(i, gstrRequestAntibody_��Ч��))), _
                        CDate(ufgData.Text(i, gstrRequestAntibody_��������)))
            
            
            '�ж��Ƿ����ʹ���˷�
            If blnIsInvalidCount Or blnIsInvalidDate Then
                If MsgBoxD(Me, "���� [" & ufgData.Text(i, gstrRequestAntibody_��������) & "]" & _
                                IIf(blnIsInvalidCount, "���޿����˷ݣ�", "") & _
                                IIf(blnIsInvalidDate, "�ѹ���Ч�ڣ�", "") & "�Ƿ������Ӹ���Ŀ��", vbYesNo, Me.Caption) <> vbYes Then
                    blnContinueAdd = False

                End If
            End If
                        
                        

            
            If blnContinueAdd Then
                strSQL = "select Zl_��������_�ؼ���Ŀ_����([1],[2],[3],[4],[5],[6],[7],[8],[9]) as ����ֵ from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                        mlngPatholAdviceId, _
                                                        GetSelectedMaterialId, _
                                                        mlngCurRequestId, _
                                                        Val(ufgData.KeyValue(i)), _
                                                        Val(cbxSpeExamType.Text), _
                                                        lngSpeexamDetails, _
                                                        IIf(chkPriceState.value <> 0, 1, 0), _
                                                        ufgData.Text(i, gstrRequestAntibody_��Ŀ˳��), _
                                                        IIf(mblnIsBuZuo, 1, 0))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveSpeExamRequest", "δ�ɹ���ȡ��������ؼ���ĿID,����ʧ�ܡ�")
                    Exit Sub
                End If
                
                '���ý�����Ϣ
                lngNewRow = mufgParentContext.NewRow
                
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_ID) = rsData!����ֵ
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_�걾����) = GetSelectSpecimenName
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_�Ŀ��) = GetSelectMaterialNum
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_��������) = ufgData.Text(i, gstrRequestAntibody_��������)
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_��������) = IIf(mblnIsBuZuo, "����", "����")
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_��ǰ״̬) = "������"
                
                Call mufgParentContext.LocateRow(lngNewRow)
                
                ufgData.RowHidden(i) = True
            End If

        End If
    Next i
    
End Sub


Private Function GetSelectedMaterialId()
'ȡ�õ�ǰѡ��ĲĿ�ID
    GetSelectedMaterialId = Substr(cbxMaterial.Text, 1, InStr(1, cbxMaterial.Text, "(") - 1)
End Function


Private Function GetSelectSpecimenName()
'ȡ�õ�ǰѡ��ı걾����
    Dim strMaterialInf As String
    Dim strReplace As String
    
    GetSelectSpecimenName = ""
    If Trim(cbxMaterial.Text) = "" Then Exit Function
    
    strMaterialInf = cbxMaterial.Text
    
    strReplace = Left(strMaterialInf, InStr(1, strMaterialInf, "�걾����") + 3)
    
    strMaterialInf = Replace(strMaterialInf, strReplace, "")
    
    
    GetSelectSpecimenName = Mid(strMaterialInf, 1, Len(strMaterialInf) - 1)
End Function



Private Function GetSelectMaterialNum()
'ȡ�õ�ǰѡ��ĲĿ��
    Dim strMaterialInf As String
    Dim strReplace As String
    
    GetSelectMaterialNum = ""
    If Trim(cbxMaterial.Text) = "" Then Exit Function
    
    strMaterialInf = cbxMaterial.Text
    
    strReplace = Mid(strMaterialInf, 1, InStr(1, strMaterialInf, "(�Ŀ�ţ�") + 4)
    
    strMaterialInf = Replace(strMaterialInf, strReplace, "")
    
    
    GetSelectMaterialNum = Mid(strMaterialInf, 1, InStr(strMaterialInf, " �걾����") - 1)
End Function

Private Sub cmdSure_Click()
'�����ؼ�����
On Error GoTo ErrHandle
    Call cmdApply_Click
    
    If blnIsOk Then Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
'��ʼ������
    mlngCurRequestId = -1
    mlngPatholAdviceId = -1
    mblnIsBuZuo = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    '��ʼ��������ʾ�б�
    Call InitAntibodyList
    
    '�����ؼ�����
    Call LoadSpeExamType
    
    '�����ؼ�ϸĿ
    Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
    
    '��ѯ������Ϣ
    Call QueryAntibodyData
    
    '�����ײ͹�������
    Call LoadMealLinkData
    
    '�����ײͷ���
    Call LoadMealClass
    
    '�����ײ���Ϣ
    Call LoadAntibodyMeal(cbxMealClass.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '�رմ���ʱ�����б�����
     zlDatabase.SetPara "�ؼ쿹���б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
    
    Set mrsMealLink = Nothing
End Sub


Private Sub LoadMealClass()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select distinct �ײ���� from �����ײ���Ϣ"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbxMealClass.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    cbxMealClass.AddItem ""
    While Not rsData.EOF
        If Nvl(rsData!�ײ����) <> "" Then cbxMealClass.AddItem Nvl(rsData!�ײ����)
        rsData.MoveNext
    Wend
End Sub

Private Sub lstMealClass_Click()
On Error GoTo ErrHandle
   
    Call LoadAntibodyDataToFace(lstMealClass.Text)
    
    Call cbxMaterial_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtAntibodyName_Change()
''���ݿ������ƹ���
'On Error GoTo errHandle
'    Dim lngFindIndex As Long
'    Dim i As Long
'    Dim strTxt() As String
'
'    If txtAntibodyName.Text = "" Then Exit Sub
'
'    strTxt() = Split(txtAntibodyName.Text, " ")
'
'    For i = LBound(strTxt) To UBound(strTxt)
'        lngFindIndex = ufgData.FindRowIndex(strTxt(i), gstrRequestAntibody_��������, True)
'        If lngFindIndex > 0 Then
'            Call ufgData.SetRowChecked(lngFindIndex, True, csSystem)
'            Call ufgData.LocateRow(lngFindIndex)
'        End If
'    Next i
'
'
''    If Trim(txtAntibodyName.Text) = "" Then Exit Sub
'
''    lngFindIndex = ufgData.FindRowIndex(txtAntibodyName.Text, gstrRequestAntibody_��������)
''
''    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'��ʾ������ϸ��Ϣ
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgData.KeyValue(lngAntibodyRow), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub txtAntibodyName_KeyPress(KeyAscii As Integer)
'���ݿ������ƹ���
On Error GoTo ErrHandle
    Dim lngFindIndex As Long
    Dim i As Long
    Dim strTxt() As String
    
    If txtAntibodyName.Text = "" Then Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    strTxt() = Split(txtAntibodyName.Text, " ")
    
    For i = LBound(strTxt) To UBound(strTxt)
        lngFindIndex = ufgData.FindRowIndex(strTxt(i), gstrRequestAntibody_��������, True)
        If lngFindIndex > 0 Then
            Call ufgData.SetRowCheck(lngFindIndex, True)
            Call ufgData.LocateRow(lngFindIndex)
        End If
    Next i
    
    
'    If Trim(txtAntibodyName.Text) = "" Then Exit Sub
    
'    lngFindIndex = ufgData.FindRowIndex(txtAntibodyName.Text, gstrRequestAntibody_��������)
'
'    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Call ShowAntibodyInf(Row)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
'˫����ʾ������ϸ
On Error GoTo ErrHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Sub
    
    Call ShowAntibodyInf(ufgData.SelectionRow)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

