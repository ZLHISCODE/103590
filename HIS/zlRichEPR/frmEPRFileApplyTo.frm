VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileApplyTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���÷�Χ"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4995
   Icon            =   "frmEPRFileApplyTo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2940
      ScaleHeight     =   315
      ScaleWidth      =   1815
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1890
      Width           =   1815
      Begin VB.ComboBox cbo����ʱ�� 
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ó���"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView lvwBakup 
      Height          =   2475
      Left            =   -510
      TabIndex        =   15
      Tag             =   "10"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3615
      TabIndex        =   14
      Top             =   5175
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2445
      TabIndex        =   13
      Top             =   5175
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -60
      TabIndex        =   12
      Top             =   5025
      Width           =   5115
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "����ʾѡ����(&L)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2910
      TabIndex        =   11
      Top             =   4785
      Width           =   1830
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -60
      TabIndex        =   10
      Top             =   585
      Width           =   5115
   End
   Begin VB.OptionButton optApply 
      Caption         =   "���������²���(&2)"
      Height          =   195
      Index           =   2
      Left            =   570
      TabIndex        =   4
      Top             =   1935
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "ȫԺͨ�ò���(&1)"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   3
      Top             =   1635
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "�ݲ�ʹ��(&0)"
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   2
      Top             =   1350
      Value           =   -1  'True
      Width           =   1950
   End
   Begin MSComctlLib.ListView lvwApply 
      Height          =   2475
      Left            =   570
      TabIndex        =   6
      Tag             =   "10"
      Top             =   2235
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫ��(&E)"
      Height          =   350
      Index           =   1
      Left            =   1650
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Index           =   0
      Left            =   570
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      Caption         =   "ʹ�÷�Χ(&S)"
      Height          =   180
      Left            =   255
      TabIndex        =   5
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�����:   001-��Ժ��¼"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   750
      Width           =   2160
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmEPRFileApplyTo.frx":058A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "���Ը���ҽѧרҵ�Ĳ�ͬҪ��ָ�����ļ������ڲ��ֲ��Ż�ȫԺͨ�á�"
      Height          =   360
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   3960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRFileApplyTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mintDef As Integer          '����
Private mstrCode As String          '�����ļ����
Private mintKind As Integer       '��������
Private mlngFileID As Long        '�����ļ�ID
Private mblnOK As Boolean
Dim objItem As ListItem


Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileId As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    mblnOK = False
    mlngFileID = lngFileId
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����, ���, ����, ͨ��,���� From �����ļ��б� Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "�ļ���ʧ(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Exit Function
        mintDef = !����
        mintKind = !����
        mstrCode = !���
        Me.lblFile.Caption = "�ļ�����:   " & !��� & "-" & !����
        Me.optApply(IIf(IsNull(!ͨ��), 0, !ͨ��)).Value = True
    End With
    
    '---------------------------------------------------
    '��ѡ��������ѡ�����б�
    With Me.lvwBakup.ColumnHeaders
        .Clear
        .Add , "_����", "����", 900
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 800
    End With
    With Me.lvwApply.ColumnHeaders
        .Clear
        .Add , "_����", "����", 900
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 800
    End With
    With Me.lvwApply
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    With Me.cbo����ʱ��
        .Clear
        .AddItem "δ�趨"
        .AddItem "����ǰ"
        .AddItem "�����"
        .ListIndex = 0
    End With

    Select Case mintKind
    Case 1
        gstrSQL = "Select d.Id, d.����, d.����, d.����, Decode(s.����id, Null, 0, 1) As ѡ��" & _
                " From ���ű� d, ��������˵�� m, (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]) s" & _
                " Where d.Id = m.����id And d.Id = s.����id(+) And m.�������� = '�ٴ�' And m.������� In (1, 3)"
    Case 2
        gstrSQL = "Select d.Id, d.����, d.����, d.����, Decode(s.����id, Null, 0, 1) As ѡ��" & _
                " From ���ű� d, ��������˵�� m, (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]) s" & _
                " Where d.Id = m.����id And d.Id = s.����id(+) And m.�������� = '�ٴ�' And m.������� In (2, 3)"
    Case 3, 4
        gstrSQL = "Select d.Id, d.����, d.����, d.����, Decode(s.����id, Null, 0, 1) As ѡ��" & _
                " From ���ű� d, ��������˵�� m, (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]) s" & _
                " Where d.Id = m.����id And d.Id = s.����id(+) And m.�������� = '����' And m.������� In (2, 3)"
    Case 5, 6
        gstrSQL = "Select d.Id, d.����, d.����, d.����, Decode(s.����id, Null, 0, 1) As ѡ��" & _
                " From ���ű� d, ��������˵�� m, (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]) s" & _
                " Where d.Id = m.����id And d.Id = s.����id(+) And m.�������� = '�ٴ�'"
    Case Else
        Unload Me: ShowMe = False: Exit Function
    End Select
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwBakup.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = "" & !����
            If !ѡ�� = 1 Then objItem.Checked = True
            
            Set objItem = Me.lvwApply.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = "" & !����
            If !ѡ�� = 1 Then objItem.Checked = True
            .MoveNext
        Loop
    End With
    
    If mintKind = 3 Then
        '����Ƿ�Ϊ����,��������Ӧ����
        Call SetObstetric
        '����ǲ���,����ȡ���ָ�����ʱ��(8)
        Dim str��ʽ As String
        str��ʽ = ";;;;;;;;;"
        gstrSQL = "Select ��ʽ From ����ҳ���ʽ Where ����=[1] And ���=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mintKind, mstrCode)
        If NVL(rsTemp!��ʽ) <> "" Then
            str��ʽ = rsTemp!��ʽ
        End If
        
        If UBound(Split(str��ʽ, ";")) >= 8 Then Me.cbo����ʱ��.ListIndex = Val(Split(str��ʽ, ";")(8))
    End If
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkSelect_Click()
Dim objAdd As ListItem
Dim objItem As ListItem

    Me.lvwApply.ListItems.Clear
    If Me.chkSelect.Value Then
        For Each objItem In Me.lvwBakup.ListItems
            If objItem.Checked Then
                Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
                objAdd.Checked = objItem.Checked
            End If
        Next
    Else
        For Each objItem In Me.lvwBakup.ListItems
            Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
            objAdd.Checked = objItem.Checked
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim arr��ʽ
    Dim strSelected As String
    Dim intStart As Integer, intEnd As Integer
    Dim str���� As String, str���� As String, str��ʽ As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡѡ��Ŀ����嵥
    strSelected = ""
    For Each objItem In Me.lvwApply.ListItems
        If objItem.Checked Then strSelected = strSelected & ";" & Mid(objItem.Key, 2)
    Next
    If strSelected <> "" Then strSelected = Mid(strSelected, 2)
    
    If Me.optApply(0).Value Then
        str���� = "Zl_�����ļ��б�_Applyto(" & mlngFileID & ",0,Null)"
    ElseIf Me.optApply(1).Value Then
        str���� = "Zl_�����ļ��б�_Applyto(" & mlngFileID & ",1,Null)"
    Else
        If strSelected = "" Then MsgBox "û��ѡ����ң�", vbInformation, gstrSysName: Me.lvwApply.SetFocus: Exit Sub
        str���� = "Zl_�����ļ��б�_Applyto(" & mlngFileID & ",2,'" & strSelected & "')"
    End If
    
    '������ƻ����¼���ķ���ʱ��
    If mintKind = 3 And mintDef <> -1 Then
        str��ʽ = ";;;;;;;"
        gstrSQL = "Select ��ʽ From ����ҳ���ʽ Where ����=[1] And ���=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mintKind, mstrCode)
        If NVL(rsTemp!��ʽ) <> "" Then
            str��ʽ = rsTemp!��ʽ
        End If
        
        'ǰ8λ����,ƴ���ϲ��Ʒ���ʱ��
        intEnd = 7
        arr��ʽ = Split(str��ʽ, ";")
        str��ʽ = ""
        For intStart = 0 To intEnd
            str��ʽ = str��ʽ & ";" & arr��ʽ(intStart)
        Next
        str��ʽ = Mid(str��ʽ, 2) & ";" & IIf(pic����.Visible, Me.cbo����ʱ��.ListIndex, 0)
        
        str���� = "Zl_����ҳ���ʽ_Format(" & mintKind & ",'" & mstrCode & "','" & str��ʽ & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    If str���� <> "" Then
        gcnOracle.BeginTrans
        blnTrans = True
    End If
    Call zldatabase.ExecuteProcedure(str����, Me.Caption)
    If str���� <> "" Then
        Call zldatabase.ExecuteProcedure(str����, Me.Caption)
        gcnOracle.CommitTrans
        blnTrans = False
    End If
    mblnOK = True
    Unload Me
    Exit Sub

errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim objItem As ListItem
    For Each objItem In Me.lvwBakup.ListItems
        objItem.Checked = IIf(Index = 0, True, False)
    Next
    Call chkSelect_Click
    Call SetObstetric
    Me.lvwApply.SetFocus
End Sub

Private Sub lvwApply_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwApply.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwApply.SortOrder = IIf(Me.lvwApply.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwApply.SortKey = ColumnHeader.Index - 1
        Me.lvwApply.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwApply_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.lvwBakup.ListItems(Item.Key).Checked = Item.Checked
    Call SetObstetric
End Sub

Private Sub lvwApply_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetObstetric
End Sub

Private Sub optApply_Click(Index As Integer)
    Me.lvwApply.Enabled = Me.optApply(2).Value
    Me.chkSelect.Enabled = Me.optApply(2).Value
    Me.cmdSelect(0).Enabled = Me.optApply(2).Value
    Me.cmdSelect(1).Enabled = Me.optApply(2).Value
    Call SetObstetric
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function IsObstetric() As Boolean
    Dim strSelected As String
    Dim intStart As Integer, intEnd As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ѡ�����б��Ƿ�Ϊ����
    
    If mintKind <> 3 Then Exit Function     'ֻ�л����ļ��Ž��������
    If mintDef = -1 Then Exit Function
    If Not optApply(2).Value Then Exit Function
    
    '��ȡѡ��Ŀ���
    strSelected = ""
    intEnd = Me.lvwApply.ListItems.Count
    For intStart = 1 To intEnd
        If lvwApply.ListItems(intStart).Checked Then
            strSelected = strSelected & "," & Mid(lvwApply.ListItems(intStart).Key, 2)
        End If
    Next
    If strSelected = "" Then Exit Function
    strSelected = Mid(strSelected, 2)
    
    '����Ƿ񶼾߱����Ƶ�����
    gstrSQL = "" & _
              " SELECT ID FROM ���ű� WHERE ID IN (Select Column_Value From Table(ZLTOOLS.f_Num2list([1])))" & vbNewLine & _
              " MINUS" & vbNewLine & _
              " SELECT ����ID FROM ��������˵�� WHERE ��������='����'"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ񶼾߱����Ƶ�����", strSelected)
    IsObstetric = (rsTemp.RecordCount = 0)  'û�зǲ��Ʋ���
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetObstetric()
    Dim blnVisible As Boolean
    '��ѡ��ʱ�ж�,�����ѡ���Ҷ����в����������������ò��Ʒ���ʱ��
    
    blnVisible = IsObstetric
    pic����.Visible = blnVisible
End Sub


