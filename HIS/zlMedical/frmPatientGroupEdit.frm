VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPatientGroupEdit 
   Caption         =   "��Ա����"
   ClientHeight    =   5580
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9645
   Icon            =   "frmPatientGroupEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9645
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2475
      Left            =   2910
      TabIndex        =   12
      Top             =   1785
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4366
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   5220
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatientGroupEdit.frx":076A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8115
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1218
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1438
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1652
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1872
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7515
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2218
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2438
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9645
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.�ָ�"
               Key             =   "�ָ�"
               Object.ToolTipText     =   "�ָ�(Alt+R)"
               Object.Tag             =   "&R.�ָ�"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   10155
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   2835
      Begin VB.CommandButton cmdClear 
         Caption         =   "��ѡ��(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   21
         Top             =   3105
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2310
         Width           =   2580
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ѡ��(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   2685
         Width           =   1470
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1665
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.����"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   7
         Top             =   2055
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.����"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   795
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.�Ա�"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.����״��"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   1425
         Width           =   900
      End
   End
   Begin VB.Frame fra1 
      Height          =   585
      Left            =   3225
      TabIndex        =   19
      Top             =   4695
      Width           =   6585
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   2580
      End
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "����(&J)"
         Height          =   350
         Left            =   3975
         TabIndex        =   15
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.����Ϊ���"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   225
         Width           =   1080
      End
   End
   Begin VB.Frame fra2 
      Height          =   540
      Left            =   2850
      TabIndex        =   20
      Top             =   1185
      Width           =   6525
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   165
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.���"
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   10
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "�ָ�(&R)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmPatientGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mstrKey As String
Private mrsMember As New ADODB.Recordset
Private mlngLoop As Long

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False
    End If
    
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("�ָ�").Enabled = mnuFileRestore.Enabled
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    vsf.AppendRow = True
    
    Call InitData
    
    EditChanged = True
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    Call ClearData
                    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function
    
    stbThis.Panels(2).Text = "�����������Ա����Ϊ��ͬ���飬���в�ͬ����졣"
    
    EditChanged = False
    
    mblnStartUp = False
    
    Call cbo_Click(3)
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
               
'    ��������Ա�����ڴ��������
    gstrSQL = "select A.����id AS ID,0 AS ѡ��,A.����,A.�����,A.�Ա�,A.����,TO_CHAR(A.��������,'yyyy-mm-dd') AS ��������,A.����״��,NVL(C.�������,'') AS ������,0 As ǰ��ɫ " & _
                "from ������Ϣ A,�����Ա���� B,(SELECT * FROM ������ WHERE �Ǽ�id=" & mlngKey & ") C  " & _
                "WHERE A.����ID=B.����ID AND B.�������=C.�������(+) AND B.�Ǽ�id=" & mlngKey & " Order By C.�������,A.�����"
    
    Set mrsMember = New ADODB.Recordset
    mrsMember.Open gstrSQL, gcnOracle, adOpenStatic, adLockBatchOptimistic
        
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    mstrKey = ""
    
    On Error GoTo errHand
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "ѡ��", 450, 1, , 1, , flexDTBoolean
        .NewColumn "����", 1200, 1
        .NewColumn "�����", 900, 7
        .NewColumn "�Ա�", 600, 1
        .NewColumn "����", 900, 1
        .NewColumn "����״��", 900, 1
        .NewColumn "��������", 1080, 1
        .NewColumn "������", 1500, 1, GetCombList("SELECT ������� FROM ������ Where �Ǽ�id=" & mlngKey), 1
        .NewColumn "", 15, 1
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .AppendRow = True
    End With
    
    cbo(2).Clear
    cbo(3).Clear
    
    cbo(3).AddItem "<����>"
    '��ȡ�����Ϣ
    gstrSQL = "SELECT ������� AS ����,ROWNUM AS ID FROM ������ WHERE �Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Call AddComboData(cbo(2), rs)
        Call AddComboData(cbo(3), rs, False)
    End If
    If cbo(2).ListCount > 0 Then cbo(2).ListIndex = 0
    If cbo(3).ListCount > 0 Then cbo(3).ListIndex = 0
    
    cbo(0).Clear
    cbo(0).AddItem "<����>"
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID FROM �Ա� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    If cbo(0).ListCount > 0 Then cbo(0).ListIndex = 0
    
    cbo(1).Clear
    cbo(1).AddItem "<����>"
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID FROM ����״�� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs, False)
    If cbo(1).ListCount > 0 Then cbo(1).ListIndex = 0
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
        
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim lngLoop As Long
    Dim rsPati As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    mrsMember.Filter = ""
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            
            strSQL = "ZL_�����Ա����_CLASS(" & mlngKey & "," & mrsMember("ID").Value & ",'" & mrsMember("������").Value & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
            
            mrsMember.MoveNext
        Loop
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    
    If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
    For lngLoop = 1 To rsSQL.RecordCount
        Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
        rsSQL.MoveNext
    Next
    
    gcnOracle.CommitTrans
    blnTran = False

    SaveEdit = True

    Exit Function

errHand:

    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans

End Function


Private Sub cbo_Click(Index As Integer)
    
    If mblnStartUp = True Then Exit Sub
    
    Select Case Index
    Case 2
        '
        
    Case 3
        
        Call ResetVsf(vsf)
        
        mrsMember.Filter = ""
        If cbo(Index).Text <> "<����>" Then
            mrsMember.Filter = "������='" & cbo(Index).Text & "'"
        End If
        
        mrsMember.Sort = "������,�����"
           
        If mrsMember.RecordCount > 0 Then
            mrsMember.MoveFirst
            Call FillGrid(vsf, mrsMember, Array("", "", "", "", "", "", "yyyy-MM-dd"))
        End If
        vsf.AppendRow = True
        
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdAdjust_Click()
    Dim lngLoop As Long
    Dim blnFlag As Boolean
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If Abs(Val(vsf.TextMatrix(lngLoop, 1))) = 1 Then
                vsf.TextMatrix(lngLoop, 8) = cbo(2).Text
                Call vsf_AfterEdit(lngLoop, 8)
                blnFlag = True
            End If
        End If
    Next
    
    If blnFlag Then Call cbo_Click(3)
    
End Sub

Private Sub cmdClear_Click()
    Dim lngLoop As Long
    Dim strFilter As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    
    strFilter = ""
    
    If cbo(3).Text <> "<����>" Then strFilter = " AND ������='" & cbo(3).Text & "'"
    If cbo(0).Text <> "<����>" Then strFilter = strFilter & " AND �Ա�='" & zlCommFun.GetNeedName(cbo(0).Text) & "'"
    If cbo(1).Text <> "<����>" Then strFilter = strFilter & " AND ����״��='" & zlCommFun.GetNeedName(cbo(1).Text) & "'"
    
    If Trim(txt(1).Text) <> "" Then strFilter = " AND ����='" & txt(1).Text & "'"
    
    varTmp2 = Split(Trim(txt(0).Text), ",")
    strTmp = ""
    For lngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(lngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR ����='" & varTmp2(lngLoop) & "'"
        Else
            strTmp = strTmp & "  OR (����>='" & Mid(varTmp2(lngLoop), 1, InStr(varTmp2(lngLoop), "-") - 1) & "' AND ����<='" & Mid(varTmp2(lngLoop), InStr(varTmp2(lngLoop), "-") + 1) & "')"
        End If
    Next
    If strTmp <> "" Then strFilter = strFilter & " AND (" & Mid(strTmp, 5) & ")"
        
    mrsMember.Filter = ""
    If strFilter <> "" Then mrsMember.Filter = Mid(strFilter, 6)
                                
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            mrsMember.Update "ѡ��", 0
            mrsMember.Update "ǰ��ɫ", 0
            
            mrsMember.MoveNext
        Loop
    End If
    
    Call cbo_Click(3)
    
'    If mrsMember.RecordCount > 0 Then
'        mrsMember.MoveFirst
'        Do While Not mrsMember.EOF
'            mrsMember.Update "ѡ��", 0
'            mrsMember.MoveNext
'        Loop
'    End If
End Sub

Private Sub cmdSelect_Click()
    
    Dim lngLoop As Long
    Dim strFilter As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    
    strFilter = ""
    
    If cbo(3).Text <> "<����>" Then strFilter = " AND ������='" & cbo(3).Text & "'"
    If cbo(0).Text <> "<����>" Then strFilter = strFilter & " AND �Ա�='" & zlCommFun.GetNeedName(cbo(0).Text) & "'"
    If cbo(1).Text <> "<����>" Then strFilter = strFilter & " AND ����״��='" & zlCommFun.GetNeedName(cbo(1).Text) & "'"
    
    If Trim(txt(1).Text) <> "" Then strFilter = " AND ����='" & txt(1).Text & "'"
    
    varTmp2 = Split(Trim(txt(0).Text), ",")
    strTmp = ""
    
    For lngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(lngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR ����='" & varTmp2(lngLoop) & "'"
        Else
            strTmp = strTmp & "  OR (����>='" & Mid(varTmp2(lngLoop), 1, InStr(varTmp2(lngLoop), "-") - 1) & "' AND ����<='" & Mid(varTmp2(lngLoop), InStr(varTmp2(lngLoop), "-") + 1) & "')"
        End If
    Next
    
'    For mlngLoop = 0 To UBound(varTmp2)
'
'        If InStr(varTmp2(mlngLoop), "-") = 0 Then
'
'            'Call GetBirth(Val(varTmp2(mlngLoop)), strStart, strEnd)
'            strTmp = strTmp & " OR (����>='" & strStart & "' AND ��������<='" & strEnd & "')"
'        Else
'
'            'Call GetBirth(Val(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1)), strStart, strEnd)
'            strTmp = strTmp & " OR (��������<='" & strEnd & "'"
'
'            'Call GetBirth(Val(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1)), strStart, strEnd)
'            strTmp = strTmp & " AND ��������>='" & strStart & "')"
'
'        End If
'    Next
    If strTmp <> "" Then strFilter = strFilter & " AND (" & Mid(strTmp, 5) & ")"
        
    mrsMember.Filter = ""
    If strFilter <> "" Then mrsMember.Filter = Mid(strFilter, 6)
                                
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            mrsMember.Update "ѡ��", 1
            mrsMember.Update "ǰ��ɫ", 16711680
            mrsMember.MoveNext
        Loop
        
    End If
    
    Call cbo_Click(3)
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyR
            If tbrThis.Buttons("�ָ�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�ָ�"))
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()
    glngFormW = 9765
    glngFormH = 6270
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With vsf
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height
        .Width = fra2.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fra1.Height + 90
    End With
    
    With fra1
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 90
        .Width = vsf.Width
    End With
    
    vsf.AppendRow = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
    
    If MsgBox("ȷʵҪ�ָ���ǰ��ѡ��Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    
    Call ReadData
        
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    
    If SaveEdit() Then
                
        On Error Resume Next
        
        EditChanged = False
        mblnOK = True
        
        Unload Me
        
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileSave_Click
    Case "�ָ�"
        Call mnuFileRestore_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 1
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 0      '
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789-,")
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 8 Then
        mrsMember.Filter = ""
        mrsMember.Filter = "ID=" & Val(vsf.RowData(Row))
        
        If mrsMember.RecordCount > 0 Then
            mrsMember("������").Value = vsf.TextMatrix(Row, Col)
            EditChanged = True
        End If
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

