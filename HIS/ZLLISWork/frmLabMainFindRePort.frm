VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmLabMainFindRePort 
   Caption         =   "���˱����ѯ"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11745
   Icon            =   "frmLabMainFindRePort.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11745
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptFind 
      Height          =   5355
      Left            =   30
      TabIndex        =   7
      Top             =   690
      Width           =   11655
      _Version        =   589884
      _ExtentX        =   20558
      _ExtentY        =   9446
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   345
      Left            =   5940
      TabIndex        =   17
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "�Ѵ�ӡ"
      Height          =   225
      Index           =   2
      Left            =   2250
      TabIndex        =   15
      Top             =   6090
      Width           =   885
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "�����"
      Height          =   225
      Index           =   1
      Left            =   1155
      TabIndex        =   14
      Top             =   6090
      Width           =   885
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "�Ѻ���"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   6090
      Width           =   885
   End
   Begin VB.CommandButton cmdUnionPrint 
      Caption         =   "�ϲ���ӡ(&U)"
      Height          =   345
      Left            =   7335
      TabIndex        =   12
      ToolTipText     =   "�Ѷ���걾�ϲ�Ϊһ�����浥���д�ӡ"
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdSetupPrint 
      Caption         =   "��ӡ����(&P)"
      Height          =   345
      Left            =   4500
      TabIndex        =   11
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   10260
      TabIndex        =   10
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   345
      Left            =   8805
      TabIndex        =   9
      Top             =   6270
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CheckBox chģ������ 
         Caption         =   "ģ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   16
         Top             =   270
         Width           =   1155
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "ʱ�䷶Χ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4290
         TabIndex        =   8
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   375
         Left            =   10170
         TabIndex        =   5
         Top             =   180
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   345
         Left            =   5550
         TabIndex        =   2
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
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
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   73138179
         CurrentDate     =   39449
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   750
         TabIndex        =   1
         Top             =   210
         Width           =   2265
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   345
         Left            =   7920
         TabIndex        =   4
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
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
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   73138179
         CurrentDate     =   39449
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   217
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         IDKindStr       =   "��|����|0;ҽ|ҽ����|1;��|���֤��|2;IC|IC����|3;��|�����|4;��|���￨|5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7650
         TabIndex        =   3
         Top             =   270
         Width           =   210
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0078
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0612
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabMainFindRePort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------- 2007-08-17 ����һ��֧ͨ��
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mintUnion As Integer
Private mstrPrivs As String
Private mbln���֤ As Boolean

Private Enum IDKinds
    C0���� = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
    C4����� = 4
    C5���￨ = 5
End Enum
Private Enum mCol
    ID
    ѡ��
    ����
    �Ա�
    ����
    �շ�״̬
    ״̬
    �����
    סԺ��
    �걾��
    ������Ŀ
    ������
    ����ʱ��
    ��������
    ��������
    ��������
    ����ҽ��
    ����ʱ��
    ������
    ����ʱ��
    ������
    ����ʱ��
    ������
    ����ʱ��
    �����
    ���ʱ��
    ҽ��id
    �걾id
End Enum
Private mblnCard As Boolean '�Ƿ�ˢ��
Private mobjSquareCard As Object                                        'ȡ������
Private mblnShowPwd As Boolean                                          '�Ƿ���ʾ����

Private Sub chkDate_Click()
    Me.DTPBegin.Enabled = Me.chkDate.Value
    Me.DTPEnd.Enabled = Me.chkDate.Value
End Sub

Private Sub chkSelect_Click(Index As Integer)
    Dim intLoop As Integer
    
    With Me.rptFind
        If .Rows.Count = 0 Then Exit Sub
        For intLoop = 0 To .Rows.Count - 1
            'ѡ���Ѻ���
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.״̬).Value = "5-�Ѻ���" Then
                    .Rows(intLoop).Record(mCol.ѡ��).Checked = (Me.chkSelect(0).Value = 1)
                End If
                'ѡ�������
                If .Rows(intLoop).Record(mCol.״̬).Value = "6-�����" Then
                    .Rows(intLoop).Record(mCol.ѡ��).Checked = (Me.chkSelect(1).Value = 1)
                End If
                'ѡ���Ѵ�ӡ
                If .Rows(intLoop).Record(mCol.״̬).Value = "7-�Ѵ�ӡ" Then
                    .Rows(intLoop).Record(mCol.ѡ��).Checked = (Me.chkSelect(2).Value = 1)
                End If
            End If
        Next
        .Redraw
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Me.cmdFind.SetFocus
    RefreshData
    Me.txtID.SetFocus
    Me.txtID.SelStart = 0
    Me.txtID.SelLength = Len(Me.txtID.Text)
End Sub

Private Sub cmdPreview_Click()
    Dim intLoop As Long
    With Me.rptFind
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow = True Then Exit Sub
        intLoop = .FocusedRow.Index
        If .Rows(intLoop).Record(mCol.״̬).Value = "5-�Ѻ���" Or .Rows(intLoop).Record(mCol.״̬).Value = "6-�����" _
            Or .Rows(intLoop).Record(mCol.״̬).Value = "7-�Ѵ�ӡ" Then
            'Ԥ��
            ReportPrint intLoop, False
        End If
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim intLoop As Integer
    Dim blnPrint As Boolean
    Dim strInfo As String
    
    With Me.rptFind
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.ѡ��).Checked = True Then
                    If .Rows(intLoop).Record(mCol.״̬).Value = "5-�Ѻ���" Or .Rows(intLoop).Record(mCol.״̬).Value = "6-�����" _
                        Or .Rows(intLoop).Record(mCol.״̬).Value = "7-�Ѵ�ӡ" Then
                        If Me.rptFind.Rows(intLoop).Record(mCol.�����).Value = "" Then
                            If InStr(mstrPrivs, "δ��˴�ӡ") <= 0 Then
                                If strInfo = "" Then
                                    MsgBox "��û��<δ��˴�ӡ>Ȩ�ޣ����ܴ�ӡδ��˵���!"
                                End If
                            Else
                                '��ӡ
                                ReportPrint intLoop, True
                            End If
                        Else
                            '��ӡ
                            ReportPrint intLoop, True
                        End If
                        
                    End If
                End If
            End If
        Next
    End With
    

End Sub

Private Sub cmdSetupPrint_Click()
    '��ӡ����
    PrintSetup
End Sub

Private Sub cmdUnionPrint_Click()
    '�ϲ���ӡ�걾
    Call AllReportPrint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyF10 Then
'        Call IdKindChange
'    End If
End Sub


Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5���￨, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    mbln���֤ = False
    
    rptFind.AllowColumnRemove = False
    rptFind.ShowItemsInGroups = False
    
    With rptFind.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
    End With
    rptFind.SetImageList ImgList
    With Me.rptFind.Columns
        Set Column = .Add(mCol.ID, "������Ϣ", 75, True): Column.Visible = False
        Set Column = .Add(mCol.ѡ��, "", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.����, "����", 60, True): Column.Visible = False
        Set Column = .Add(mCol.�Ա�, "�Ա�", 60, True): Column.Visible = False
        Set Column = .Add(mCol.����, "����", 60, True): Column.Visible = False
        Set Column = .Add(mCol.�շ�״̬, "�շ�״̬", 75, True): Column.Visible = False
        Set Column = .Add(mCol.״̬, "״̬", 100, True): Column.Visible = False
        Set Column = .Add(mCol.�����, "�����", 100, True): Column.Visible = False
        Set Column = .Add(mCol.סԺ��, "סԺ��", 100, True): Column.Visible = False
        Set Column = .Add(mCol.�걾��, "�걾��", 60, True)
        Set Column = .Add(mCol.������Ŀ, "������Ŀ", 100, True)
        Set Column = .Add(mCol.������, "������", 100, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.��������, "��������", 100, True)
        Set Column = .Add(mCol.��������, "��������", 100, True)
        Set Column = .Add(mCol.��������, "��������", 100, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.������, "������", 100, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.������, "������", 100, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.������, "������", 100, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.�����, "�����", 100, True)
        Set Column = .Add(mCol.���ʱ��, "���ʱ��", 100, True)
        Set Column = .Add(mCol.ҽ��id, "ҽ��ID", 100, True)
        Set Column = .Add(mCol.�걾id, "�걾ID", 100, True)
    End With
    
    Me.DTPEnd.Value = Now
    Me.DTPBegin.Value = Now - 30
    Me.chkDate.Value = zlDatabase.GetPara("frmLabMainFindRePort_ʹ��ʱ�䷶Χ", 100, 1208, 0)
    Me.rptFind.LoadSettings zlDatabase.GetPara("frmLabMainFindRePort_rptFind", 100, 1208, "")
    mintUnion = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0)
    Me.DTPBegin.Enabled = Me.chkDate.Value
    Me.DTPEnd.Enabled = Me.chkDate.Value
    IDKind.IDKind = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���뷽ʽ", 0))
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind��ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
    
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
End Sub

Private Sub RefreshData()
    '����       ˢ������
    Dim strWhere As String
    Dim strFind As String
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim GroupRow As ReportRow
    Dim blnBarCode As Boolean
    Dim strSQLbak As String
    Dim lng�����ID As Long
    Dim lng����ID As Long
    
    If Trim(Me.txtID.Text) = "" Then Me.txtID.SetFocus: Exit Sub
    
    If mbln���֤ Or IDKind.IDKind = IDKind.GetKindIndex("���֤��") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtID, False, lng����ID) = False Then lng����ID = 0
        txtID = "-" & lng����ID
'        strSQL = "select ����ID from ������Ϣ where ���֤�� = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
'        If Not rsTmp.EOF Then
'            txtID = "-" & rsTmp.Fields("����ID")
'        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("IC����") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtID, False, lng����ID) = False Then lng����ID = 0
        txtID = "-" & lng����ID
'        strSQL = "select ����ID from ������Ϣ where IC���� = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
'        If Not rsTmp.EOF Then
'            txtID = "-" & rsTmp.Fields("����ID")
'        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("���￨") Then
        strSQL = "select ����ID from ������Ϣ where ���￨�� = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
        If Not rsTmp.EOF Then
            txtID.Tag = txtID.Text
            txtID = "-" & rsTmp.Fields("����ID")
        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtID, False, lng����ID) = False Then lng����ID = 0
        txtID = "-" & lng����ID
'        If InStr("-+*./", Mid(Me.txtID.Text, 1, 1)) <= 0 Then
'            Me.txtID.Text = "*" & Me.txtID.Text
'        End If
    Else
        If Val(IDKind.GetKindItem("�����ID")) <> 0 Then
            lng�����ID = Val(IDKind.GetKindItem("�����ID"))
            If mobjSquareCard.zlGetPatiID(lng�����ID, txtID, False, lng����ID) = False Then lng����ID = 0
            If lng����ID = 0 Then lng����ID = 0
        Else
            lng����ID = 0
'            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtID, False, lng����ID) = False Then lng����ID = 0
        End If
        If lng����ID > 0 Then
            txtID.Tag = txtID.Text
            txtID = "-" & lng����ID
        End If
    End If
    
    Select Case Mid(Me.txtID, 1, 1)
        Case "-"                                '����ID
            strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ����id = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "+"                                'סԺ��
            strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where סԺ�� = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "*"                                '�����
            strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ����� = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "."                                '�Һŵ���
            strWhere = "(select b.����ID,b.����,b.�Ա�,b.����,b.�����,b.סԺ�� from ����ҽ����¼ a , ������Ϣ b ��where  a.����id = b.����ID and �Һŵ� = [1] ) e "
            strFind = Mid(Me.txtID, 2)
        Case "/"                                '�շѵ��ݺ�
            strWhere = "(select  b.����ID,b.����,b.�Ա�,b.����,b.�����,b.סԺ�� from ������ü�¼ a, ������Ϣ b " & _
                       " where No = [1] and a.����id = b.����id ) e "
            strFind = zlCommFun.GetFullNO(Mid(txtID, 2))
        Case Else                               '���￨������
            strFind = Me.txtID
            If IDKind.IDKind = IDKind.GetKindIndex("����") And BlnIsNumber(strFind) Then
                    strWhere = "( select C.* from ����ҽ����¼ a , ����ҽ������ b , ������Ϣ C " & _
                         " Where a.ID = b.ҽ��id And a.����ID = C.����ID and  b.�������� = [1] ) e "
                    blnBarCode = True
            Else
                If mblnCard Or IDKind.IDKind = IDKind.GetKindIndex("���￨") Then
                    strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ���￨�� = [1]) e "
                    strFind = UCase(Me.txtID)
                ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Then
                    strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ����� = [1]) e "
                    strFind = Val(Mid(Me.txtID, 2))
                Else
                    If chģ������.Value = 1 Then
                        strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ���� || ''  like '%' || [1] || '%' ) e "
                        chkDate.Enabled = True
                    ElseIf Len(Me.txtID.Text) = 1 Then
                        strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ���� || '' like  [1] || '%' ) e "
                        chkDate.Enabled = True
                    Else
                        strWhere = "(select " & gConst_������Ϣ_���� & " from ������Ϣ a where ���� like   [1] || '%' ) e "
                    End If
                End If
            End If
    End Select
    mblnCard = False
    
    gstrSql = "select /*+ rule */ distinct a.����id, a.����, a.�Ա�, a.����, a.�����, סԺ��, ������Ŀ, ״̬, ������, ����ʱ��, ��������, ��������," & vbNewLine & _
                "       ��������, ����ҽ��, ����ʱ��,������, ����ʱ��, ������, ����ʱ��, ������, ����ʱ��," & vbNewLine & _
                "       a.�����, a.���ʱ��, ҽ��ID, �걾ID," & vbNewLine & _
                "       b.��¼����, b.��¼״̬, b.�����־,Ӥ��,����1,�Ա�1, �걾���,�걾�����ʾ��" & vbNewLine & _
                "from (select " & vbNewLine & _
                "       distinct e.����id, e.����, e.�Ա�, e.����, e.�����, e.סԺ��, a.ҽ������ as ������Ŀ," & vbNewLine & _
                "                decode(b.ҽ��id, null, '1-δ����', decode(b.������, Null, '2-δ����', decode(b.������, null, '3-�Ѳ���', '4-�ѽ���'))) as ״̬," & vbNewLine & _
                "                b.������, b.����ʱ��, b.��������," & vbNewLine & _
                "                decode(b.��¼����, 1, '�շ�', 2, '����') as ��������, d.���� as ��������, a.����ҽ��," & vbNewLine & _
                "                a.����ʱ��, b.������, b.����ʱ��, b.������, b.����ʱ��, '' as ������, '' as ����ʱ��," & vbNewLine & _
                "                '' as �����, '' as ���ʱ��, a.id as ҽ��Id, '' as �걾ID, b.��¼����,a.Ӥ��,e.�Ա� as �Ա�1,e.���� as ����1, " & vbNewLine & _
                "                '' as �걾�����ʾ, '' as �걾���,a.id as ���ҽ��ID " & vbNewLine & _
                "       from ����ҽ����¼ a, ����ҽ������ b, ���ű� d " & "," & strWhere & vbNewLine & _
                "       where a.id = b.ҽ��id(+) and a.��������id = d.id and a.����id = e.����id and" & vbNewLine & _
                "             b.ִ��״̬ = 0 and a.������� = 'C' and a.������Դ = 2 " & vbNewLine & _
                " " & IIf(blnBarCode = True, " and b.�������� = [4] ", " ") & vbNewLine & _
                " " & IIf(Me.chkDate.Value = 0, "", " and a.����ʱ�� between [2] and [3] ") & vbNewLine

    gstrSql = gstrSql & "       union all " & vbNewLine & _
                "       select " & vbNewLine & _
                "       distinct e.����id, e.����, e.�Ա�, e.����, e.�����, e.סԺ��, a.������Ŀ," & vbNewLine & _
                "" & vbNewLine & _
                "                decode(����״̬, 1, '5-�Ѻ���', decode(sign(nvl(��ӡ����, 0)), 1, '7-�Ѵ�ӡ', '6-�����')) as ״̬," & vbNewLine & _
                "                d.������, d.����ʱ��, d.��������," & vbNewLine & _
                "                decode(d.��¼����, 1, '�շ�', 2, '����') as ��������, f.���� as ��������, b.����ҽ��," & vbNewLine & _
                "                b.����ʱ��, d.������, d.����ʱ��, d.������, d.����ʱ��, a.������," & vbNewLine & _
                "                to_char(a.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��, a.�����, to_char(a.���ʱ��,'YYYY-MM-DD HH24:MI:SS') as ���ʱ��," & vbNewLine & _
                "                a.ҽ��ID, to_char(a.ID) as �걾ID, d.��¼����,b.Ӥ��,a.�Ա� as �Ա�1,a.���� as ����1, " & vbNewLine & _
                "                Decode(a.����id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                "                 a.�걾���) As �걾�����ʾ, a.�걾���,b.id ���ҽ��ID " & vbNewLine & _
                "       from ����걾��¼ a, ����ҽ����¼ b, ����ҽ������ d, ���ű� f " & "," & strWhere & vbNewLine & _
                "       where a.ҽ��id = b.���id and b.���id = d.ҽ��id and b.��������ID = f.id and" & vbNewLine & _
                "             a.����id = e.����id And a.������Դ = 2 " & vbNewLine & _
                " " & IIf(blnBarCode = True, " And D.�������� = [4] ", " ") & vbNewLine & _
                " " & IIf(Me.chkDate.Value = 0, "", " and a.����ʱ�� between [2] and [3] ") & vbNewLine & _
                ") a, סԺ���ü�¼ b" & vbNewLine & _
                "where a.���ҽ��ID = b.ҽ�����(+) and a.��¼���� = b.��¼����(+) and nvl(b.��¼״̬,0) in (0,1) "

    strSQLbak = gstrSql
    strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
    strSQLbak = Replace$(strSQLbak, "a.������Դ = 2", "a.������Դ <> 2")
    gstrSql = gstrSql & " union  " & strSQLbak & " order by ����id, ״̬, ����ʱ��, ����ʱ�� "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strFind, CDate(Format(Me.DTPBegin.Value, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DTPEnd.Value, "yyyy-mm-dd 23:59:59")), strFind)
    If rsTmp.RecordCount > 0 Then Me.txtID.Text = "": Me.txtID.SetFocus
    Me.rptFind.Records.DeleteAll
    Me.rptFind.GroupsOrder.DeleteAll
    
    On Error GoTo errH
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڲ���������ȴ�....", Me
    
    Do Until rsTmp.EOF
        With Me.rptFind
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
        End With
        If Nvl(rsTmp("״̬")) = "5-�Ѻ���" Or Nvl(rsTmp("״̬")) = "6-�����" Or Nvl(rsTmp("״̬")) = "7-�Ѵ�ӡ" Then
            Record.Item(mCol.ѡ��).HasCheckbox = True
            If blnBarCode = True Then
                Record.Item(mCol.ѡ��).Checked = True
            End If
        Else
            Record.Item(mCol.ѡ��).HasCheckbox = False
        End If
        Record.Item(mCol.ID).Value = Nvl(rsTmp("����id"))
        Record.Item(mCol.����).Value = Nvl(rsTmp("����")) & IIf(Nvl(rsTmp("Ӥ��"), 0) > 0, "(Ӥ��" & rsTmp("Ӥ��") & ")", "")
        Record.Item(mCol.�Ա�).Value = IIf(Nvl(rsTmp("Ӥ��"), 0) = 0, Nvl(rsTmp("�Ա�")), Nvl(rsTmp("�Ա�1")))
        Record.Item(mCol.����).Value = IIf(Nvl(rsTmp("Ӥ��"), 0) = 0, Nvl(rsTmp("����")), Nvl(rsTmp("����1")))
        '������Ϣ
        Record.Item(mCol.ID).GroupCaption = "����:" & Record.Item(mCol.����).Value & " �Ա�:" & Record.Item(mCol.�Ա�).Value & " ����:" & _
                                    Replace(Nvl(Record.Item(mCol.����).Value), "Ӥ��", "") & _
                                    " �����:" & Nvl(rsTmp("�����")) & " סԺ��" & Nvl(rsTmp("סԺ��"))
        Record.Item(mCol.�����).Value = Nvl(rsTmp("�����"))
        Record.Item(mCol.סԺ��).Value = Nvl(rsTmp("סԺ��"))
        Record.Item(mCol.������Ŀ).Value = Nvl(rsTmp("������Ŀ"))
        '�շ�״̬
        Select Case Nvl(rsTmp("��¼״̬"))
            Case ""     'δ�շ�
                Record.Item(mCol.�շ�״̬).Value = "δ�շ�"
            Case "0"    '���۵�
                If Nvl(rsTmp("�����־")) = 1 Then
                    Record.Item(mCol.�շ�״̬).Value = "����" & Nvl(rsTmp("��������")) & "(���۵�)"
                Else
                    Record.Item(mCol.�շ�״̬).Value = "סԺ" & Nvl(rsTmp("��������")) & "(���۵�)"
                End If
            Case "1"    '���ʺ��շ����
                If Nvl(rsTmp("�����־")) = 1 Then
                    Record.Item(mCol.�շ�״̬).Value = "����" & Nvl(rsTmp("��������")) & IIf(Nvl(rsTmp("��������")) = "�շ�", "(���շ�)", "(�Ѽ���)")
                Else
                    Record.Item(mCol.�շ�״̬).Value = "סԺ" & Nvl(rsTmp("��������")) & IIf(Nvl(rsTmp("��������")) = "�շ�", "(���շ�)", "(�Ѽ���)")
                End If
        End Select
        Record.Item(mCol.״̬).Value = Nvl(rsTmp("״̬"))
        Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
        Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
        Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
        Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
        Record.Item(mCol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
        Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
        Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
        Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
        Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
        Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.�����).Value = Nvl(rsTmp("�����"))
        Record.Item(mCol.���ʱ��).Value = Format(Nvl(rsTmp("���ʱ��")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.ҽ��id).Value = Nvl(rsTmp("ҽ��ID"))
        Record.Item(mCol.�걾id).Value = Nvl(rsTmp("�걾ID"))
        Record.Item(mCol.�걾��).Value = Nvl(rsTmp("�걾���"))
        Record.Item(mCol.�걾��).Caption = Nvl(rsTmp("�걾�����ʾ"))
        rsTmp.MoveNext
    Loop
    For intLoop = 0 To Me.rptFind.Columns.Count - 1
        If Me.rptFind.Columns(intLoop).Caption = "������Ϣ" Then
            Me.rptFind.GroupsOrder.Add Me.rptFind.Columns(intLoop)
            Me.rptFind.GroupsOrder(0).Editable = True
        End If
    Next
    For intLoop = 0 To Me.rptFind.Columns.Count - 1
        If Me.rptFind.Columns(intLoop).Caption = "״̬" Then
            Me.rptFind.GroupsOrder.Add Me.rptFind.Columns(intLoop)
            Me.rptFind.GroupsOrder(1).Editable = True
        End If
    
    Next
    Me.rptFind.Populate
    '
    For Each GroupRow In rptFind.Rows
        If GroupRow.GroupRow = False Then
            If GroupRow.Record(mCol.״̬).Value = "1-δ����" Or GroupRow.Record(mCol.״̬).Value = "2-δ����" Or _
                GroupRow.Record(mCol.״̬).Value = "3-�Ѳ���" Or GroupRow.Record(mCol.״̬).Value = "4-�ѽ���" Then
                GroupRow.ParentRow.Expanded = False
            End If
        End If
    Next
    
    zlCommFun.StopFlash
    
    If IDKind.IDKind = IDKinds.C5���￨ And Me.txtID.Tag <> "" Then
        txtID = txtID.Tag
        txtID.Tag = ""
    End If
    
    Me.MousePointer = 0
    
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.Frame1
        .Left = 10
        .Width = Me.ScaleWidth - 10
    End With
    
    With Me.cmdExit
        .Top = Me.ScaleHeight - 500
        .Left = Me.ScaleWidth - .Width - 300
    End With
    
    With Me.cmdPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdExit.Left - .Width - 400
    End With
    
    With Me.cmdUnionPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdPrint.Left - .Width - 400
    End With
    
    With Me.cmdPreview
        .Top = Me.cmdExit.Top
        .Left = Me.cmdUnionPrint.Left - .Width - 400
    End With
    
    With Me.cmdSetupPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdPreview.Left - .Width - 400
    End With
    
    With Me.rptFind
        .Width = Me.ScaleWidth
        .Height = Me.cmdExit.Top - .Top - 150
    End With
    
    Me.chkSelect(0).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    Me.chkSelect(1).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    Me.chkSelect(2).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    zlDatabase.SetPara "frmLabMainFindRePort_ʹ��ʱ�䷶Χ", Me.chkDate.Value, 100, 1208
    zlDatabase.SetPara "frmLabMainFindRePort_rptFind", Me.rptFind.SaveSettings, 100, 1208
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���뷽ʽ", IDKind.IDKind)

End Sub

Private Sub IDKind_Click()
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
    If IDKind.IDKind = IDKind.GetKindIndex("IC����") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtID.Text = mobjICCard.Read_Card()
            If txtID.Text <> "" Then Call RefreshData
        End If
    End If
    lng�����ID = Val(IDKind.GetKindItem("�����ID"))
    If lng�����ID = 0 Then Exit Sub
    
    If mobjSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtID.Text = strOutCardNO
    If txtID.Text <> "" Then Call txtID_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txtID = ""
    If mblnShowPwd = True Then
        Me.txtID.PasswordChar = "*"
    Else
        Me.txtID.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    mbln���֤ = False
    If Not txtID.Locked And txtID.Text = "" And Me.ActiveControl Is txtID Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2���֤��
        txtID.Text = strID
        mbln���֤ = True
        Call RefreshData
        mbln���֤ = False
        IDKind.IDKind = lngPreIDKind
    End If
End Sub

Private Sub rptFind_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error Resume Next
    If Item.Record(mCol.״̬).Value = "5-�Ѻ���" Or Item.Record(mCol.״̬).Value = "6-�����" Or Item.Record(mCol.״̬).Value = "7-�Ѵ�ӡ" Then
        Item.Record(mCol.ѡ��).Checked = Not Item.Record(mCol.ѡ��).Checked
        Me.rptFind.Redraw
    End If
End Sub

Private Sub rptFind_SelectionChanged()
    If Me.rptFind.FocusedRow Is Nothing Then Me.cmdPrint.Enabled = False: Me.cmdSetupPrint.Enabled = False: Exit Sub
    If Me.rptFind.FocusedRow.GroupRow = True Then Me.cmdPrint.Enabled = False: Me.cmdSetupPrint.Enabled = False: Exit Sub

    If Me.rptFind.FocusedRow.Record(mCol.״̬).Value = "5-�Ѻ���" Or Me.rptFind.FocusedRow.Record(mCol.״̬).Value = "6-�����" _
       Or Me.rptFind.FocusedRow.Record(mCol.״̬).Value = "7-�Ѵ�ӡ" Then
        Me.cmdPrint.Enabled = True
    Else
        Me.cmdPrint.Enabled = False
    End If
    Me.cmdSetupPrint.Enabled = True
End Sub

Private Sub txtID_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtID.Text = "" And Me.ActiveControl Is txtID)
End Sub

Private Sub txtID_GotFocus()
    If Not mobjIDCard Is Nothing And txtID.Text = "" And Not txtID.Locked Then mobjIDCard.SetEnabled (True)
    txtID.SelStart = 0
    txtID.SelLength = Len(txtID.Text)
    txtID.SetFocus
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,��.��""") = True Then KeyAscii = 0
    blnCard = False
    If IDKind.IDKind = IDKind.GetKindIndex("����") Then
'        mblnCard = zlCommFun.InputIsCard(txtID, KeyAscii, False)
        If mblnCard = False And KeyAscii = 13 Then
            KeyAscii = 0
            cmdFind_Click
        End If
    End If
    If IDKind.IDKind = IDKind.GetKindIndex("���￨") Then
'        Call zlCommFun.InputIsCard(txtID, KeyAscii, True)
        gbytCardNOLen = Val(IDKind.GetKindItem("���ų���", IDKind.IDKind))
        blnCard = KeyAscii <> 8 And Len(txtID.Text) = gbytCardNOLen - 1 And txtID.SelLength <> Len(txtID.Text)
        If blnCard = True Then
            If KeyAscii <> 13 Then
                Me.txtID = Me.txtID & Chr(KeyAscii)
            End If
            KeyAscii = 0
            cmdFind_Click
        End If
    End If
    If KeyAscii = 13 Or (IDKind.IDKind = IDKind.GetKindIndex("���￨") And blnCard = True) Then
        Call cmdFind_Click
    End If
End Sub

Private Sub txtID_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub
Private Sub ReportPrint(ByVal intIndex As Integer, ByVal blnPrint As Boolean)
    '���������ӡ
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڴ�ӡ��ȴ�...", Me
    
    If Me.rptFind.Rows(intIndex) Is Nothing Then
        zlCommFun.StopFlash
        Me.MousePointer = 0
        Exit Sub
    End If
    
    lngҽ��ID = Val(Me.rptFind.Rows(intIndex).Record(mCol.ҽ��id).Value)
    lng����ID = Val(Me.rptFind.Rows(intIndex).Record(mCol.ID).Value)
    lngKey = Val(Me.rptFind.Rows(intIndex).Record(mCol.�걾id).Value)
    
    '����ͼ�ι��Զ��屨�����
    strSQL = "select id from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngKey)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                        "����ID=" & lng����ID, "�걾ID=" & lngKey, "���ҽ��=" & lngҽ��ID, "����걾=" & lngKey, _
                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                        "ͼ��9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH
    If blnPrint = True Then
        If Me.rptFind.Rows(intIndex).Record(mCol.״̬).Value = "6-�����" Then
            If mintUnion = 1 Then
                gstrSql = " select id from ����걾��¼ where ҽ��id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҽ��ID)
                Do Until rsTmp.EOF
                    strSQL = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
                    zlDatabase.ExecuteProcedure strSQL, gstrSysName
                    rsTmp.MoveNext
                Loop
            Else
                strSQL = "ZL_����걾��¼_�걾�ʿ�(" & lngKey & ",'',1)"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            End If
        End If
    End If
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(lngKey As Long, Objfrm As Object, strPrivs As String)
    '��ʾ���˲��˱��洰��
    mstrPrivs = strPrivs
    Me.Show , Objfrm
    Me.txtID.Text = "-" & lngKey
    Call cmdFind_Click
    Me.txtID.Text = ""
End Sub
Private Sub PrintSetup()
    '��ӡ����
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    Dim intLoop  As Integer
    
   
    If Me.rptFind.FocusedRow Is Nothing Then
        MsgBox "ѡ��һ����¼��������ã�", vbInformation, Me.Caption
        Exit Sub
    End If
    lngҽ��ID = Val(rptFind.FocusedRow.Record(mCol.ҽ��id).Value)
'    lng����id = Val(rptList.FocusedRow.Record(mCol.����ID).Value)
'
'    strsql = "select ���ͺ� from ����ҽ������ a , ����ҽ����¼ b where b.id = a.ҽ��id and b.id = [1]"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, gstrSysName, lngҽ��ID)
'    If rsTmp.EOF = False Then
'        lng���ͺ� = Nvl(rsTmp(0))
'    End If
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
        
    End If
End Sub
Private Sub AllReportPrint()
    '���������ӡ
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    Dim strҽ��ID As String                 'ҽ��ID�����ҽ��IDʹ��","�ָ���
    Dim str�걾ID As String                 '�걾ID, ����걾IDʹ��","�ָ���
    Dim strPrintCode As String              '���ݱ���
    Dim intItem As Integer
    Dim astrItem() As String
        
    
    If Me.rptFind.Rows.Count = 0 Then Exit Sub
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڴ�ӡ��ȴ�...", Me
    For intLoop = 0 To Me.rptFind.Rows.Count - 1
        If Me.rptFind.Rows(intLoop).GroupRow = False Then
            If Me.rptFind.Rows(intLoop).Record(mCol.ѡ��).Checked = True Then
                If Me.rptFind.Rows(intLoop).Record(mCol.�����).Value = "" Then
                    If InStr(mstrPrivs, "δ��˴�ӡ") <= 0 Then
                        MsgBox "��û��<δ��˴�ӡ>Ȩ�ޣ����ܴ�ӡδ��˵���!"
                        Me.MousePointer = 1
                        zlCommFun.StopFlash
                        Exit Sub
                    End If
                End If
                strҽ��ID = strҽ��ID & "," & Me.rptFind.Rows(intLoop).Record(mCol.ҽ��id).Value
                str�걾ID = str�걾ID & "," & Me.rptFind.Rows(intLoop).Record(mCol.�걾id).Value
                lng����ID = Me.rptFind.Rows(intLoop).Record(mCol.ID).Value
            End If
        End If
    Next
    If strҽ��ID <> "" Then
        strҽ��ID = Mid(strҽ��ID, 2)
        lngҽ��ID = Split(strҽ��ID, ",")(0)
    End If
    If str�걾ID <> "" Then
        str�걾ID = Mid(str�걾ID, 2)
        lngKey = Split(str�걾ID, ",")(0)
    End If
    
    '�ж����ʽʱ�õ���ʽ
    frmLabMainPrintFormat.ShowMe Me, strҽ��ID, strPrintCode
    
    '����ͼ�ι��Զ��屨�����
    astrItem = Split(Mid(str�걾ID, 2), ",")
    intLoop = 1
    For intItem = 0 To UBound(astrItem)
        If intLoop >= 9 Then Exit For
        frmLabMain.ReadImageData CLng(astrItem(intItem)), True
        strSQL = "select id from ����ͼ���� where �걾id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, astrItem(intItem))
        Do Until rsTmp.EOF
            If intLoop < 9 Then
                strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                intLoop = intLoop + 1
            End If
            rsTmp.MoveNext
        Loop
    Next
    
    Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & strҽ��ID, _
                        "����ID=" & lng����ID, "�걾ID=" & str�걾ID, "���ҽ��=" & strҽ��ID, "����걾=" & str�걾ID, _
                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                        "ͼ��9=" & strChart(9), 2)


    On Error GoTo errH
    For intLoop = 0 To Me.rptFind.Rows.Count - 1
        If Me.rptFind.Rows(intLoop).GroupRow = False Then
            If Me.rptFind.Rows(intLoop).Record(mCol.״̬).Value = "6-�����" Then
                If mintUnion = 1 Then
                    gstrSql = " select id from ����걾��¼ where ҽ��id = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҽ��ID)
                    Do Until rsTmp.EOF
                        strSQL = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
                        zlDatabase.ExecuteProcedure strSQL, gstrSysName
                        rsTmp.MoveNext
                    Loop
                Else
                    strSQL = "ZL_����걾��¼_�걾�ʿ�(" & lngKey & ",'',1)"
                    zlDatabase.ExecuteProcedure strSQL, gstrSysName
                End If
            End If
        End If
    Next
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
