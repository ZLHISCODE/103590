VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmElementChange 
   Caption         =   "����Ҫ����������"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   Icon            =   "frmElementChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7050
   ScaleMode       =   0  'User
   ScaleWidth      =   14453.51
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   6975
      Left            =   4815
      TabIndex        =   28
      ToolTipText     =   "˫���б��������п������������޸�"
      Top             =   15
      Width           =   9525
      _Version        =   589884
      _ExtentX        =   16801
      _ExtentY        =   12303
      _StockProps     =   0
      BorderStyle     =   2
      ShowGroupBox    =   -1  'True
      ShowItemsInGroups=   -1  'True
   End
   Begin VB.Frame fraThis 
      Height          =   7065
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   4740
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3405
         TabIndex        =   2
         Top             =   6510
         Width           =   1080
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   2340
         TabIndex        =   1
         Top             =   6510
         Width           =   1080
      End
      Begin VB.Frame fraBecouse 
         Caption         =   "�䶯ԭ��"
         Height          =   2310
         Left            =   210
         TabIndex        =   3
         Top             =   225
         Width           =   4290
         Begin VB.CommandButton cmdDisease 
            Caption         =   "��"
            Height          =   225
            Left            =   3765
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   1470
            Width           =   240
         End
         Begin VB.CommandButton cmdDiagnose 
            Caption         =   "��"
            Height          =   225
            Left            =   3765
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   1845
            Width           =   240
         End
         Begin VB.TextBox txtDiagnose 
            Height          =   270
            Left            =   1485
            TabIndex        =   12
            Top             =   1815
            Width           =   2550
         End
         Begin VB.TextBox txtDisease 
            Height          =   270
            Left            =   1485
            TabIndex        =   11
            Top             =   1440
            Width           =   2550
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "Ҫ��"
            Height          =   225
            Index           =   0
            Left            =   405
            TabIndex        =   10
            Top             =   405
            Value           =   -1  'True
            Width           =   690
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "���˼���"
            Height          =   225
            Index           =   1
            Left            =   390
            TabIndex        =   9
            Top             =   1470
            Width           =   1035
         End
         Begin VB.ComboBox cboElName 
            Height          =   300
            Left            =   1710
            TabIndex        =   7
            Text            =   "cboElName"
            Top             =   360
            Width           =   2325
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "�������"
            Height          =   225
            Index           =   2
            Left            =   390
            TabIndex        =   5
            Top             =   1830
            Width           =   1035
         End
         Begin VB.ComboBox cboContent 
            Height          =   300
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   915
            Width           =   2325
         End
         Begin VB.Label lblElName 
            Caption         =   "����"
            Height          =   225
            Left            =   1215
            TabIndex        =   14
            Top             =   405
            Width           =   435
         End
         Begin VB.Label lblElVal 
            Caption         =   "����"
            Height          =   225
            Left            =   1215
            TabIndex        =   13
            Top             =   945
            Width           =   435
         End
      End
      Begin VB.Frame fraSo 
         Caption         =   "�䶯���"
         Height          =   3555
         Left            =   210
         TabIndex        =   15
         Top             =   2715
         Width           =   4290
         Begin VB.ComboBox cboAddSentence 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   3180
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "׷�Ӵʾ�"
            Height          =   225
            Index           =   4
            Left            =   210
            TabIndex        =   34
            Top             =   3218
            Width           =   1185
         End
         Begin VB.ComboBox cboSameElName 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2715
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "��ͬҪ��ͬʱ���"
            Height          =   360
            Index           =   3
            Left            =   210
            TabIndex        =   31
            Top             =   2685
            Width           =   1185
         End
         Begin VB.ComboBox cboDelElname 
            Height          =   300
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2235
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "ɾ��Ҫ��"
            Height          =   225
            Index           =   2
            Left            =   225
            TabIndex        =   29
            Top             =   2280
            Width           =   1260
         End
         Begin VB.OptionButton optSo 
            Caption         =   "Ҫ��"
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   21
            Top             =   345
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.ComboBox cboSoElname 
            Height          =   300
            Left            =   1695
            TabIndex        =   20
            Text            =   "cboSoElname"
            Top             =   300
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "�ʾ�"
            Height          =   225
            Index           =   1
            Left            =   225
            TabIndex        =   19
            Top             =   1200
            Width           =   705
         End
         Begin VB.TextBox txtSoElContent 
            Height          =   270
            Left            =   1695
            TabIndex        =   18
            ToolTipText     =   "�Էֺŷָ�"
            Top             =   780
            Width           =   2325
         End
         Begin VB.ComboBox cboSoStCompend 
            Height          =   300
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1425
            Width           =   2325
         End
         Begin VB.ComboBox cboSoSentence 
            Height          =   300
            Left            =   1695
            TabIndex        =   16
            Text            =   "cboSoSentence"
            Top             =   1755
            Width           =   2325
         End
         Begin VB.Label lblSoStCompend 
            Caption         =   "�������"
            Height          =   225
            Left            =   825
            TabIndex        =   26
            Top             =   1470
            Width           =   810
         End
         Begin VB.Label lblSoElname 
            Caption         =   "����"
            Height          =   225
            Left            =   1200
            TabIndex        =   25
            Top             =   345
            Width           =   435
         End
         Begin VB.Label lblSoElContent 
            Caption         =   "����(��"";""�ָ�)"
            Height          =   225
            Left            =   285
            TabIndex        =   24
            Top             =   810
            Width           =   1350
         End
         Begin VB.Label Label5 
            Caption         =   "����ʹ�õĴʾ�"
            Height          =   225
            Left            =   1200
            TabIndex        =   23
            Top             =   1200
            Width           =   1350
         End
         Begin VB.Label lblSentence 
            Caption         =   "�ʾ�����"
            Height          =   225
            Left            =   825
            TabIndex        =   22
            Top             =   1785
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   1275
         TabIndex        =   27
         Top             =   6510
         Width           =   1080
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&D)"
         Height          =   350
         Left            =   210
         TabIndex        =   33
         Top             =   6510
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmElementChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng�ļ�ID As Long
Private Enum mCol
    ID = 0
    ԭ������
    ԭ��
    �������
    ���
    ԭ��Ҫ��ID
    ԭ��Ҫ������
    ԭ��Ҫ������
    ԭ����ID
    ԭ��������
    ������id
    ����������
    ���Ҫ��ID
    ���Ҫ������
    ���Ҫ��ֵ��
    ���ԭʼֵ��
    ����ʾ�ID
    ����ʾ�����
    ���ñ���
End Enum
Public Sub ShowMe(ByVal objParent As Object, ByVal lng�ļ�ID As Long)
    mlng�ļ�ID = lng�ļ�ID
    Me.Show vbModal, objParent
End Sub
Private Sub cboAddSentence_KeyPress(KeyAscii As Integer)
Call zlControl.CboSetIndex(cboAddSentence.hWnd, zlControl.CboMatchIndex(cboAddSentence.hWnd, KeyAscii))
End Sub

Private Sub cboDelElname_KeyPress(KeyAscii As Integer)
Call zlControl.CboSetIndex(cboDelElname.hWnd, zlControl.CboMatchIndex(cboDelElname.hWnd, KeyAscii))
End Sub

Private Sub cboElName_Click()
Dim i As Integer, strItems As String, strItem As String
    On Error Resume Next '�����������Ϊ�ڵ��ô����õ�
    cboContent.Clear
    strItems = Split(lblElName.Tag, "|")(cboElName.ListIndex)
    For i = 0 To UBound(Split(strItems, ";"))
        strItem = Trim(Split(strItems, ";")(i))
        If strItem <> "�Զ���" Then
            cboContent.AddItem strItem
        End If
    Next
    If cboContent.ListCount > 0 Then cboContent.ListIndex = 0
End Sub

Private Sub cboElName_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboElName.hWnd, zlControl.CboMatchIndex(cboElName.hWnd, KeyAscii))
    If KeyAscii = vbKeyReturn Then cboElName_Click
End Sub

Private Sub cboSameElName_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSameElName.hWnd, zlControl.CboMatchIndex(cboSameElName.hWnd, KeyAscii))
End Sub

Private Sub cboSoElname_Click()
Dim i As Integer, strTmp As String
    On Error Resume Next '�����������Ϊ�ڵ��ô����õ�
    txtSoElContent.Text = ""
    txtSoElContent.Tag = Split(lblSoElname.Tag, "|")(cboSoElname.ListIndex)
    For i = 0 To UBound(Split(txtSoElContent.Tag, ";"))
        strTmp = strTmp & ";" & Trim(Split(txtSoElContent.Tag, ";")(i))
    Next
    txtSoElContent.Text = Mid(strTmp, 2)
End Sub
Private Sub cboSoElname_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSoElname.hWnd, zlControl.CboMatchIndex(cboSoElname.hWnd, KeyAscii))
    If KeyAscii = vbKeyReturn Then cboSoElname_Click
End Sub

Private Sub cboSoSentence_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSoSentence.hWnd, zlControl.CboMatchIndex(cboSoSentence.hWnd, KeyAscii))
End Sub

Private Sub cboSoStCompend_Click()
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select b.Id, b.����,zlSpellCode(b.����) ���� From ������ٴʾ� A, �����ʾ�ʾ�� B Where a.���id = [1] And a.�ʾ����id = b.����id Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ٴʾ�", Val(Split(cboSoStCompend.Tag, ";")(cboSoStCompend.ListIndex)))
    cboSoSentence.Clear: cboSoSentence.Tag = ""
    Do Until rsTemp.EOF
        cboSoSentence.Tag = cboSoSentence.Tag & rsTemp!ID & ";"
        cboSoSentence.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop
End Sub

Private Function Validate() As Boolean
Dim i As Integer
    If optBecause(0).Value Then
        If cboElName.Text = "" Then
            MsgBox "�䶯ԭ���Ҫ�ز���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
        If cboContent.Text = "" Then
            MsgBox "�䶯ԭ���Ҫ��ѡ���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If

    ElseIf optBecause(1).Value Then
        If Trim(txtDisease.Text) = "" Or Val(txtDisease.Tag) = 0 Then
            MsgBox "�䶯ԭ�򼲲�����Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optBecause(2).Value Then
        If Trim(txtDiagnose.Text) = "" Or Val(txtDiagnose.Tag) = 0 Then
            MsgBox "�䶯ԭ����ϲ���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf fraBecouse.Enabled Then
        MsgBox "û��ָ���䶯ԭ�����飡", vbInformation, gstrSysName
        Exit Function
    End If

    If optSo(0).Value Then
        If cboSoElname.Text = "" Then
            MsgBox "�䶯�����Ҫ�ز���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
        If txtSoElContent.Text = "" Then
            MsgBox "�䶯�����Ҫ��ѡ���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If

        For i = 0 To UBound(Split(txtSoElContent.Text, ";"))
            If InStr(txtSoElContent.Tag, Trim(Split(txtSoElContent.Text, ";")(i))) < 1 Then
                MsgBox "�䶯���Ҫ��ѡ���ԭ��ѡ���У�", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    ElseIf optSo(1).Value Then
        If cboSoSentence.Text = "" Then
            MsgBox "�䶯����ʾ䲻��Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(2).Value Then
        If cboDelElname.Text = "" Then
            MsgBox "�䶯���Ϊɾ��Ҫ��ʱ��ɾ����Ҫ�ز���Ϊ�գ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(3).Value Then
        If cboSameElName.Text = "" Then
            MsgBox "�䶯���Ϊ��ͬҪ��ͬʱ���ʱ�����Ҫ�ز���Ϊ��", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(4).Value Then
        If cboAddSentence.Text = "" Then
            MsgBox "�䶯���Ϊ׷�Ӵʾ�ʱ��׷�Ӵʾ䲻��Ϊ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    If optBecause(0).Value And optSo(0).Value Then
        If CLng(Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))) = CLng(Val(Split(cboSoElname.Tag, ";")(cboSoElname.ListIndex))) _
            And zl9ComLib.zlStr.NeedName(cboElName.Text) = zl9ComLib.zlStr.NeedName(cboSoElname.Text) Then
            MsgBox "�䶯ԭ��Ҫ�ز�����������䶯��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If optBecause(0).Value And optSo(2).Value Then
        If CLng(Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))) = CLng(Val(Split(cboDelElname.Tag, ";")(cboDelElname.ListIndex))) _
            And zl9ComLib.zlStr.NeedName(cboElName.Text) = zl9ComLib.zlStr.NeedName(cboDelElname.Text) Then
            MsgBox "�䶯ԭ��Ҫ�ز�����������ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Validate = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    On Error GoTo errHand
    If MsgBox("ȷʵҪɾ����ǰ�������õ�����������ϵ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "Zl_�����䶯���_Delete(" & mlng�ļ�ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "ɾ��ԭ����"
    
    Call InitList
    Call FillVfgList
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Exit Sub
    End If
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHand
    With rptList
        If .FocusedRow Is Nothing Then MsgBox "����ѡ����Ҫɾ�����С�", vbInformation, gstrSysName: Exit Sub
        If .FocusedRow.GroupRow Then MsgBox "����ѡ����Ҫɾ�����С�", vbInformation, gstrSysName: Exit Sub '������
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Sub
        gstrSQL = "Zl_�����䶯���_Delete(" & mlng�ļ�ID & "," & .FocusedRow.Record.Item(mCol.ID).Value & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "ɾ��ԭ����"
    End With
    Call InitList
    Call FillVfgList
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Exit Sub
    End If
End Sub

Private Sub cmdDiagnose_Click()
    txtDiagnose.Text = ""
    Call txtDiagnose_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdDisease_Click()
    txtDisease.Text = ""
    Call txtDisease_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdOK_Click()
'������������(����/Ҫ��)����ͬһҪ��
    On Error GoTo errHand
    If Not Validate Then Exit Sub
    If cmdOK.Caption = "�޸�(&M)" Then '˫���б�֮���ʾ�޸�,��Ҫ��ɾ��ԭ�м�¼
        If MsgBox("��ȷʵҪ�޸ĵ�ǰѡ�е���������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then cmdOK.Caption = "����(&A)": Exit Sub
        gstrSQL = "Zl_�����䶯���_Delete(" & mlng�ļ�ID & "," & rptList.FocusedRow.Record.Item(mCol.ID).Value & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "ɾ��ԭ����"
    End If
    
    Dim intBecause   As Integer, lngBecauseID As Long, strElname As String, strElContent As String
    Dim intSo As Integer, lngSoCompendID As Long, lngSoID As Long, strSoElname As String, strSoElContent As String, strSoElOldContent As String
        
    If Not fraBecouse.Enabled Then '��ͬҪ��ͬʱ���
        intBecause = 4
    ElseIf optBecause(0).Value Then 'Ҫ��
        intBecause = 1
    ElseIf optBecause(1).Value Then '����
        intBecause = 2
    ElseIf optBecause(2).Value Then '���
        intBecause = 3
    End If
    
    Select Case intBecause
        Case 1
            lngBecauseID = Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))
            strElname = zl9ComLib.zlStr.NeedName(cboElName.Text)
            strElContent = cboContent.Text
        Case 2
            lngBecauseID = Val(txtDisease.Tag)
            strElname = ""
            strElContent = ""
        Case 3
            lngBecauseID = Val(txtDiagnose.Tag)
            strElname = ""
            strElContent = ""
        Case 4
            lngBecauseID = Val(Split(cboSameElName.Tag, ";")(cboSameElName.ListIndex))
            strElname = zl9ComLib.zlStr.NeedName(cboSameElName.Text)
            strElContent = ""
    End Select
    
    If optSo(0).Value Then
        intSo = 1
    ElseIf optSo(1).Value Then
        intSo = 2
    ElseIf optSo(2).Value Then
        intSo = 3
    ElseIf optSo(3).Value Then
        intSo = 4
    ElseIf optSo(4).Value Then
        intSo = 5
    End If
    
    Select Case intSo
        Case 1
            lngSoCompendID = 0
            lngSoID = Val(Split(cboSoElname.Tag, ";")(cboSoElname.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboSoElname.Text)
            strSoElContent = txtSoElContent.Text
            strSoElOldContent = Split(lblSoElname.Tag, "|")(cboSoElname.ListIndex)
        Case 2
            lngSoCompendID = Val(Split(cboSoStCompend.Tag, ";")(cboSoStCompend.ListIndex))
            lngSoID = Val(Split(cboSoSentence.Tag, ";")(cboSoSentence.ListIndex))
            strSoElname = ""
            strSoElContent = ""
            strSoElOldContent = ""
        Case 3
            lngSoCompendID = 0
            lngSoID = Val(Split(cboDelElname.Tag, ";")(cboDelElname.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboDelElname.Text)
            strSoElContent = ""
            strSoElOldContent = ""
        Case 4
            lngSoCompendID = 0
            lngSoID = Val(Split(cboSameElName.Tag, ";")(cboSameElName.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboSameElName.Text)
            strSoElContent = ""
            strSoElOldContent = ""
        Case 5
            lngSoCompendID = 0
            lngSoID = Val(Split(cboAddSentence.Tag, ";")(cboAddSentence.ListIndex))
            strSoElname = ""
            strSoElContent = ""
            strSoElOldContent = ""
    End Select
        
    gstrSQL = "Zl_�����䶯���_Update(" & mlng�ļ�ID & "," & intBecause & "," & lngBecauseID & ",'" & strElname & "','" & strElContent & "'," & _
                    intSo & "," & lngSoCompendID & "," & lngSoID & ",'" & strSoElname & "','" & strSoElContent & "','" & strSoElOldContent & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "����ԭ����"
    
    Call InitList
    Call FillVfgList
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitList
    Call FillVfgList
End Sub
Private Sub InitList()
Dim rptCol As ReportColumn

    With rptList
        .Columns.DeleteAll
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ������, "ԭ������", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ��, "ԭ��", 80, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Sortable = False: rptCol.Visible = True
        Set rptCol = .Columns.Add(mCol.�������, "�������", 0, False): rptCol.Editable = False: rptCol.Groupable = False:: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 80, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Sortable = True: rptCol.Visible = True
        Set rptCol = .Columns.Add(mCol.ԭ��Ҫ��ID, "ԭ��Ҫ��ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ��Ҫ������, "ԭ��Ҫ��", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ��Ҫ������, "ԭ��ѡ��", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ����ID, "ԭ����ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ԭ��������, "ԭ��������", 100, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.������id, "������id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����������, "������", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Ҫ��ID, "���Ҫ��ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Ҫ������, "���Ҫ��", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Ҫ��ֵ��, "���ֵ��", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���ԭʼֵ��, "���ԭʼֵ��", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ʾ�ID, "����ʾ�ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ʾ�����, "����ʾ�", 140, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.���ñ���, "���ñ���", 1280, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = True
         
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = True
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = ""
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
End Sub
Private Sub FillVfgList()
Dim rsBecouse As ADODB.Recordset, rsSo As ADODB.Recordset
Dim rptRcd As ReportRecord, i As Integer, rptItem As ReportRecordItem, strShowF As String, strShowS As String

    rptList.Records.DeleteAll
    gstrSQL = "Select ID, ԭ��, ԭ��Ҫ��id, ԭ��Ҫ������, ԭ��Ҫ������, ԭ����id, ԭ��������" & vbNewLine & _
                "From (Select a.Id, 1 ԭ��, b.����Ҫ��id ԭ��Ҫ��id, b.Ҫ������ ԭ��Ҫ������, a.ԭ������ ԭ��Ҫ������, 0 ԭ����id, '' ԭ��������" & vbNewLine & _
                "       From �����䶯ԭ�� A, �����ļ��ṹ B" & vbNewLine & _
                "       Where a.�����ļ�id = [1] And a.�䶯ԭ�� = 1 And a.�����ļ�id = b.�ļ�id And a.ԭ��Ҫ��id = Nvl(b.����Ҫ��id, 0) And a.ԭ��Ҫ�� = b.Ҫ������" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 2 ԭ��, 0 ԭ��Ҫ��id, '' ԭ��Ҫ������, '' ԭ��Ҫ������, b.Id ԭ����id, b.���� ԭ��������" & vbNewLine & _
                "       From �����䶯ԭ�� A, ��������Ŀ¼ B" & vbNewLine & _
                "       Where a.�����ļ�id = [1] And a.�䶯ԭ�� = 2 And a.ԭ��Ҫ��id = b.Id" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 3 ԭ��, 0 ԭ��Ҫ��id, '' ԭ��Ҫ������, '' ԭ��Ҫ������, b.Id ԭ����id, b.���� ԭ��������" & vbNewLine & _
                "       From �����䶯ԭ�� A, �������Ŀ¼ B" & vbNewLine & _
                "       Where a.�����ļ�id = [1] And a.�䶯ԭ�� = 3 And a.ԭ��Ҫ��id = b.Id" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 4 ԭ��, b.����Ҫ��id ԭ��Ҫ��id, b.Ҫ������ ԭ��Ҫ������, a.ԭ������ ԭ��Ҫ������, 0 ԭ����id, '' ԭ��������" & vbNewLine & _
                "       From �����䶯ԭ�� A, �����ļ��ṹ B" & vbNewLine & _
                "       Where a.�����ļ�id = [1] And a.�䶯ԭ�� = 4 And a.�����ļ�id = b.�ļ�id And a.ԭ��Ҫ��id = Nvl(b.����Ҫ��id, 0) And a.ԭ��Ҫ�� = b.Ҫ������)"
    Set rsBecouse = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�䶯ԭ��", mlng�ļ�ID)
    Do Until rsBecouse.EOF
        gstrSQL = "Select ���, ������id, ����������, ���Ҫ��id, ���Ҫ������, ���Ҫ��ֵ��,���ԭʼֵ��, ����ʾ�id, ����ʾ�����" & vbNewLine & _
                    "From (Select 1 ���, 0 ������id, '' ����������, b.����Ҫ��id ���Ҫ��id, b.Ҫ������ ���Ҫ������, a.���ֵ�� ���Ҫ��ֵ��,a.ԭʼֵ�� ���ԭʼֵ��, 0 ����ʾ�id, '' ����ʾ�����" & vbNewLine & _
                    "       From �����䶯��� A, �����ļ��ṹ B" & vbNewLine & _
                    "       Where a.�䶯ԭ��id = [1] And a.�䶯��� = 1 And b.�ļ�id = [2] And a.���Ҫ��id = Nvl(b.����Ҫ��id, 0) And a.���Ҫ�� = b.Ҫ������" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 2 ���, b.Id ������id, b.�����ı� ����������, 0 ���Ҫ��id, '' ���Ҫ������, '' ���Ҫ��ֵ��,'' ���ԭʼֵ�� , c.Id ����ʾ�id, c.���� ����ʾ�����" & vbNewLine & _
                    "       From �����䶯��� A, �����ļ��ṹ B, �����ʾ�ʾ�� C" & vbNewLine & _
                    "       Where a.�䶯ԭ��id = [1] And a.�䶯��� = 2 And b.�ļ�id = [2] And a.�������id = b.Id And a.���Ҫ��id = c.Id" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 3 ���, 0 ������id, '' ����������, b.����Ҫ��id ���Ҫ��id, b.Ҫ������ ���Ҫ������, a.���ֵ�� ���Ҫ��ֵ��,a.ԭʼֵ�� ���ԭʼֵ��, 0 ����ʾ�id, '' ����ʾ�����" & vbNewLine & _
                    "       From �����䶯��� A, �����ļ��ṹ B" & vbNewLine & _
                    "       Where a.�䶯ԭ��id = [1] And a.�䶯��� = 3 And b.�ļ�id = [2] And a.���Ҫ��id = Nvl(b.����Ҫ��id, 0) And a.���Ҫ�� = b.Ҫ������" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 4 ���, 0 ������id, '' ����������, b.����Ҫ��id ���Ҫ��id, b.Ҫ������ ���Ҫ������, a.���ֵ�� ���Ҫ��ֵ��,a.ԭʼֵ�� ���ԭʼֵ��, 0 ����ʾ�id, '' ����ʾ�����" & vbNewLine & _
                    "       From �����䶯��� A, �����ļ��ṹ B" & vbNewLine & _
                    "       Where a.�䶯ԭ��id = [1] And a.�䶯��� = 4 And b.�ļ�id = [2] And a.���Ҫ��id = Nvl(b.����Ҫ��id, 0) And a.���Ҫ�� = b.Ҫ������" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 5 ���, 0 ������id, '' ����������, 0 ���Ҫ��id, '' ���Ҫ������, '' ���Ҫ��ֵ��,'' ���ԭʼֵ�� , c.Id ����ʾ�id, c.���� ����ʾ�����" & vbNewLine & _
                    "       From �����䶯��� A, �����ʾ�ʾ�� C" & vbNewLine & _
                    "       Where a.�䶯ԭ��id = [1] And a.�䶯��� = 5 And a.���Ҫ��id = c.Id)"
        Set rsSo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�䶯���", CLng(rsBecouse!ID), mlng�ļ�ID)
        Do Until rsSo.EOF
            strShowF = "": strShowS = ""
            With rptList
                Set rptRcd = rptList.Records.Add()
                rptRcd.AddItem CStr(Val(rsBecouse!ID))
                rptRcd.AddItem CStr(Val(rsBecouse!ԭ��))
                Select Case Val(rsBecouse!ԭ��)
                    Case 1
                        Set rptItem = rptRcd.AddItem("Ҫ�ر��")
                        rptItem.GroupCaption = CStr("��ΪҪ��ѡ���������ı仯")
                        strShowF = "��Ҫ��<" & NVL(rsBecouse!ԭ��Ҫ������) & ">" & " ѡ��:" & (NVL(rsBecouse!ԭ��Ҫ������)) & " ��ѡ�к�,"
                    Case 2
                        Set rptItem = rptRcd.AddItem("���ϼ���")
                        rptItem.GroupCaption = CStr("��Ϊ<������>���ϡ�������׼����������ı仯")
                        strShowF = "������������Ϊ:<" & NVL(rsBecouse!ԭ��������) & ">ʱ,"
                    Case 3
                        Set rptItem = rptRcd.AddItem("�������")
                        rptItem.GroupCaption = CStr("��Ϊ<������>���ϡ���ϱ�׼����������ı仯")
                        strShowF = "������������Ϊ:<" & NVL(rsBecouse!ԭ��������) & ">ʱ,"
                    Case 4
                        Set rptItem = rptRcd.AddItem("��ͬҪ��")
                        rptItem.GroupCaption = CStr("��Ϊ����<��ͬҪ��ͬʱ����>��������ı仯")
                        strShowF = "��Ҫ��<" & NVL(rsBecouse!ԭ��Ҫ������) & ">" & "���ݷ����仯ʱ������������ͬҪ��ͬʱ����"
                    Case Else: rptRcd.AddItem CStr("δ֪����")
                End Select
                rptRcd.AddItem CStr(Val(rsSo!���))
                Select Case Val(rsSo!���)
                    Case 1
                        rptRcd.AddItem CStr("Ҫ�ر仯")
                        strShowS = "Ҫ��<" & NVL(rsSo!���Ҫ������) & ">�Ŀ�ѡ��Ϊ:(" & NVL(rsSo!���Ҫ��ֵ��) & ")"
                    Case 2
                        rptRcd.AddItem CStr("����ʾ�")
                        strShowS = "����ʾ�:<" & NVL(rsSo!����ʾ�����) & ">�����:<" & NVL(rsSo!����������) & ">�г���"
                    Case 3:  rptRcd.AddItem CStr("���Ҫ��")
                        strShowS = "ɾ�������ڵ�Ҫ��<" & NVL(rsSo!���Ҫ������) & ">"
                    Case 4:  rptRcd.AddItem CStr("Ҫ�ظ���")
                        strShowS = ""
                    Case 5: rptRcd.AddItem CStr("׷�Ӵʾ�")
                        strShowS = "�Զ�׷�Ӵʾ�<" & rsSo!����ʾ����� & ">�ڵ�ǰλ��"
                    Case Else: rptRcd.AddItem CStr("δ֪����")
                End Select

                rptRcd.AddItem CStr(NVL(rsBecouse("ԭ��Ҫ��id"), 0))
                rptRcd.AddItem CStr(NVL(rsBecouse!ԭ��Ҫ������))
                rptRcd.AddItem CStr(NVL(rsBecouse!ԭ��Ҫ������))
                rptRcd.AddItem CStr(NVL(rsBecouse("ԭ����ID"), 0))
                rptRcd.AddItem CStr(NVL(rsBecouse!ԭ��������))
                rptRcd.AddItem CStr(NVL(rsSo("������ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!����������))
                rptRcd.AddItem CStr(NVL(rsSo("���Ҫ��ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!���Ҫ������))
                rptRcd.AddItem CStr(NVL(rsSo!���Ҫ��ֵ��))
                rptRcd.AddItem CStr(NVL(rsSo!���ԭʼֵ��))
                rptRcd.AddItem CStr(NVL(rsSo("����ʾ�ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!����ʾ�����))
                rptRcd.AddItem strShowF & strShowS
            End With
            rsSo.MoveNext
        Loop
        rsBecouse.MoveNext
    Loop
    
    rptList.GroupsOrder.Add rptList.Columns.Find(mCol.ԭ��)
    rptList.GroupsOrder(0).SortAscending = True
    rptList.Populate
    
    If Me.Visible = False Then
        Call optBecause_Click(0)
        Call optSo_Click(0)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rptList.Move fraThis.Width + 50, 0, Me.ScaleWidth - fraThis.Width - 50, Me.ScaleHeight
    
End Sub

Private Sub optBecause_Click(Index As Integer)
    txtDisease.Enabled = False: cmdDisease.Enabled = False
    txtDiagnose.Enabled = False: cmdDiagnose.Enabled = False
    cboElName.Enabled = False: cboElName.Clear
    cboContent.Enabled = False: cboContent.Clear
    Select Case Index
        Case 0
            cboElName.Enabled = True: cboElName.Clear
            cboContent.Enabled = True: cboContent.Clear
            Dim rsTemp As ADODB.Recordset
            gstrSQL = "Select ����Ҫ��id,zlSpellCode(Ҫ������) Ҫ�ؼ���, Ҫ������, Ҫ��ֵ��,��������" & vbNewLine & _
                        "From �����ļ��ṹ" & vbNewLine & _
                        "Where �ļ�id = [1] And �������� = 4 And Ҫ�ر�ʾ In (2, 3)" & vbNewLine & _
                        "Order By Ҫ������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪ��", mlng�ļ�ID)
            cboElName.Clear: cboElName.Tag = ""
            Do Until rsTemp.EOF
                cboElName.Tag = cboElName.Tag & NVL(rsTemp!����Ҫ��ID, 0) & ";" '��¼����Ҫ��ID���ԡ�;���ָ�,ά����cboelName.listindexͬ��
                cboElName.AddItem Replace(rsTemp!Ҫ�ؼ���, "-", "") & "-" & rsTemp!Ҫ������
                lblElName.Tag = lblElName.Tag & rsTemp!Ҫ��ֵ�� & "|"   '��"|"�ָ�,ά����cboelName.listindexͬ��,ÿ��ֵ����";"�ָ�
                rsTemp.MoveNext
            Loop
            If cboElName.ListCount > 0 Then cboElName.ListIndex = 0
        Case 1
            txtDisease.Enabled = True: cmdDisease.Enabled = True
        Case 2
            txtDiagnose.Enabled = True: cmdDiagnose.Enabled = True
    End Select
End Sub

Private Sub optSo_Click(Index As Integer)
Dim rsTemp As ADODB.Recordset
    fraBecouse.Enabled = True
    optBecause(1).Enabled = True
    optBecause(2).Enabled = True
    
    cboDelElname.Enabled = False: cboDelElname.Clear
    cboSameElName.Enabled = False: cboSameElName.Clear
    cboSoStCompend.Enabled = False: cboSoStCompend.Clear
    cboSoSentence.Enabled = False: cboSoSentence.Clear
    cboSoElname.Enabled = False: cboSoElname.Clear
    cboAddSentence.Enabled = False: cboAddSentence.Clear
    txtSoElContent.Enabled = False: txtSoElContent.Text = ""
    
    Select Case Index
        Case 0
            cboSoElname.Enabled = True: cboSoElname.Clear
            txtSoElContent.Enabled = True: txtSoElContent.Text = ""
            gstrSQL = "Select ����Ҫ��id,zlSpellCode(Ҫ������) Ҫ�ؼ���, Ҫ������,Ҫ��ֵ��,��������" & vbNewLine & _
                        "From �����ļ��ṹ" & vbNewLine & _
                        "Where �ļ�id = [1] And �������� = 4 And Ҫ�ر�ʾ In (2, 3)" & vbNewLine & _
                        "Order By Ҫ������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪ��", mlng�ļ�ID)
            cboSoElname.Clear: cboSoElname.Tag = "": lblSoElname.Tag = ""
            Do Until rsTemp.EOF
                If Val(Mid(NVL(rsTemp!��������, ""), 5, 1)) = 0 Then '��̬���������������
                    cboSoElname.Tag = cboSoElname.Tag & NVL(rsTemp!����Ҫ��ID, 0) & ";" '��¼����Ҫ��ID���ԡ�;���ָ�,ά����cboelName.listindexͬ��
                    cboSoElname.AddItem Replace(rsTemp!Ҫ�ؼ���, "-", "") & "-" & rsTemp!Ҫ������
                    lblSoElname.Tag = lblSoElname.Tag & rsTemp!Ҫ��ֵ�� & "|"   '��"|"�ָ�,ά����cboelName.listindexͬ��,ÿ��ֵ����";"�ָ�
                End If
                rsTemp.MoveNext
            Loop
            If cboSoElname.ListCount > 0 Then cboSoElname.ListIndex = 0
        Case 1
            cboSoStCompend.Enabled = True: cboSoStCompend.Clear
            cboSoSentence.Enabled = True: cboSoSentence.Clear
            gstrSQL = "Select ID,�����ı� �������" & vbNewLine & _
                        "From �����ļ��ṹ" & vbNewLine & _
                        "   Where �ļ�id = [1] And �������� = 1" & vbNewLine & _
                        "Order By ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���", mlng�ļ�ID)
            cboSoStCompend.Clear: cboSoStCompend.Tag = ""
            Do Until rsTemp.EOF
                cboSoStCompend.Tag = cboSoStCompend.Tag & rsTemp!ID & ";"
                cboSoStCompend.AddItem rsTemp!�������
                rsTemp.MoveNext
            Loop
            If cboSoStCompend.ListCount > 0 Then cboSoStCompend.ListIndex = 0
        Case 2
            cboDelElname.Enabled = True: cboDelElname.Clear
            gstrSQL = "Select ����Ҫ��id,zlSpellCode(Ҫ������) Ҫ�ؼ���, Ҫ������,Ҫ��ֵ��,��������" & vbNewLine & _
                        "From �����ļ��ṹ" & vbNewLine & _
                        "Where �ļ�id = [1] And �������� = 4" & vbNewLine & _
                        "Order By Ҫ������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪ��", mlng�ļ�ID)
            cboDelElname.Clear: cboDelElname.Tag = ""
            Do Until rsTemp.EOF
                    cboDelElname.Tag = cboDelElname.Tag & NVL(rsTemp!����Ҫ��ID, 0) & ";" '��¼����Ҫ��ID���ԡ�;���ָ�,ά����cboelName.listindexͬ��
                    cboDelElname.AddItem Replace(rsTemp!Ҫ�ؼ���, "-", "") & "-" & rsTemp!Ҫ������
                rsTemp.MoveNext
            Loop
            If cboDelElname.ListCount > 0 Then cboDelElname.ListIndex = 0
        Case 3
            fraBecouse.Enabled = False
            cboSameElName.Enabled = True: cboSameElName.Clear
            gstrSQL = "Select ����Ҫ��id,zlSpellCode(Ҫ������) Ҫ�ؼ���, Ҫ������,Ҫ��ֵ��,��������" & vbNewLine & _
                        "From �����ļ��ṹ" & vbNewLine & _
                        "Where �ļ�id = [1] And �������� = 4" & vbNewLine & _
                        "Order By Ҫ������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪ��", mlng�ļ�ID)
            cboSameElName.Clear: cboSameElName.Tag = ""
            Do Until rsTemp.EOF
                cboSameElName.Tag = cboSameElName.Tag & NVL(rsTemp!����Ҫ��ID, 0) & ";" '��¼����Ҫ��ID���ԡ�;���ָ�,ά����cboelName.listindexͬ��
                cboSameElName.AddItem Replace(rsTemp!Ҫ�ؼ���, "-", "") & "-" & rsTemp!Ҫ������
                rsTemp.MoveNext
            Loop
            If cboSameElName.ListCount > 0 Then cboSameElName.ListIndex = 0
        Case 4
            optBecause(0).Enabled = True
            optBecause(1).Enabled = False
            optBecause(2).Enabled = False
            optBecause(0).Value = True
            cboAddSentence.Enabled = True: cboAddSentence.Clear
            gstrSQL = "Select c.Id,zlSpellCode(C.����) ����,C.����" & vbNewLine & _
                        "From �����ļ��ṹ A, ������ٴʾ� B, �����ʾ�ʾ�� C" & vbNewLine & _
                        "Where a.�ļ�id = [1] And a.�������� = 1 And a.Id = b.���id And b.�ʾ����id =C.����ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ôʾ�", mlng�ļ�ID)
            cboAddSentence.Clear: cboAddSentence.Tag = ""
            Do Until rsTemp.EOF
                cboAddSentence.Tag = cboAddSentence.Tag & NVL(rsTemp!ID, 0) & ";"
                cboAddSentence.AddItem Replace(rsTemp!����, "-", "") & "-" & rsTemp!����
                rsTemp.MoveNext
            Loop
            If cboAddSentence.ListCount Then cboAddSentence.ListIndex = 0
    End Select

End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    With rptList
        If .FocusedRow Is Nothing Or .FocusedRow.GroupRow Then Exit Sub '������
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Sub
        cmdOK.Caption = "�޸�(&M)"
    End With
End Sub
Private Function ShowItem() As Boolean
    On Error GoTo errHand
    With rptList
        If .FocusedRow Is Nothing Then Exit Function
        If .FocusedRow.GroupRow Then Exit Function '������
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Function
        
        With .FocusedRow.Record
            fraBecouse.Enabled = True
            '����ԭ����ʾ
            Select Case .Item(mCol.ԭ������).Value
                Case 1
                    optBecause(0).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboElName, .Item(mCol.ԭ��Ҫ������).Value)
                    Call zl9ComLib.cbo.SeekIndex(cboContent, .Item(mCol.ԭ��Ҫ������).Value)
                Case 2
                    optBecause(1).Value = True
                    txtDisease.Tag = .Item(mCol.ԭ����ID).Value
                    txtDisease.Text = .Item(mCol.ԭ��������).Value
                Case 3
                    optBecause(2).Value = True
                    txtDiagnose.Tag = .Item(mCol.ԭ����ID).Value
                    txtDiagnose.Text = .Item(mCol.ԭ��������).Value
                Case 4
                    optBecause(0).Value = False: optBecause(1).Value = False: optBecause(2).Value = False
                    fraBecouse.Enabled = False
            End Select
            
            '��������ʾ
            Select Case .Item(mCol.�������).Value
                Case 1
                    optSo(0).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboSoElname, .Item(mCol.���Ҫ������).Value)
                    txtSoElContent.Text = .Item(mCol.���Ҫ��ֵ��).Value
                Case 2
                    optSo(1).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboSoStCompend, .Item(mCol.����������).Value)
                    Call zl9ComLib.cbo.SeekIndex(cboSoSentence, .Item(mCol.����ʾ�����).Value)
                Case 3
                    optSo(2).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboDelElname, .Item(mCol.���Ҫ������).Value)
                Case 4
                    optSo(3).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboDelElname, .Item(mCol.���Ҫ������).Value)
                Case 5
                    optSo(4).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboAddSentence, .Item(mCol.����ʾ�����).Value)
            End Select
        End With
    End With
    If Err.Number <> 0 Then GoTo errHand
    Exit Function
errHand:
    MsgBox "��ǰ�����ļ����ݷ����仯������������ԭ�����ָ�������Ѳ����ڣ������Զ�ɾ����", vbInformation, gstrSysName
    Call cmdDel_Click
    Err.Number = 0: Err.Clear
End Function
Private Sub rptList_SelectionChanged()
    Call ShowItem
    If cmdOK.Caption = "�޸�(&M)" Then
        optBecause(0).Value = True
        optSo(0).Value = True
        cmdOK.Caption = "����(&A)"
    End If
End Sub

Private Sub txtDiagnose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim objDiagnose As RECT, rsTemp As ADODB.Recordset
        objDiagnose = GetControlRect(txtDiagnose.hWnd)
        If Trim(txtDiagnose.Text) <> "" Then
            gstrSQL = "Select A.ID,A.����,A.���� From �������Ŀ¼ A,������ϱ��� B Where A.ID=B.���ID AND (A.����=[1]" & _
                                                            " or " & ZLCommFun.GetLike("A", "����", txtDiagnose.Text) & _
                                                            " or " & ZLCommFun.GetLike("B", "����", txtDiagnose.Text) & ")" & _
                                                            " And (A.����ʱ�� Is Null Or A.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Else
            gstrSQL = "Select ID,����, ���� From �������Ŀ¼ A Where ����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')"
        End If
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ѡ�񼲲�", True, txtDiagnose.Text, "", True, True, True, objDiagnose.Left, objDiagnose.Top, txtDiagnose.Height, True, False, True, CStr(txtDiagnose.Text))
        If Not rsTemp Is Nothing Then
            txtDiagnose.Tag = rsTemp!ID: txtDiagnose.Text = rsTemp!����
        Else
            zlControl.TxtSelAll txtDisease
        End If
    End If
End Sub

Private Sub txtDisease_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim objDisease As RECT, rsTemp As ADODB.Recordset
        objDisease = GetControlRect(txtDisease.hWnd)
        If Trim(txtDisease.Text) <> "" Then
            gstrSQL = "Select ID,����, ���� From ��������Ŀ¼ A Where (����=[1]" & _
                                                            " or " & ZLCommFun.GetLike("A", "����", txtDisease.Text) & _
                                                            " or " & ZLCommFun.GetLike("A", "����", txtDisease.Text) & _
                                                            " or " & ZLCommFun.GetLike("A", "�����", txtDisease.Text) & ")" & _
                                                            " And (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Else
            gstrSQL = "Select ID,����, ���� From ��������Ŀ¼ A Where ����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')"
        End If
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ѡ�񼲲�", True, txtDisease.Text, "", True, True, True, objDisease.Left, objDisease.Top, txtDisease.Height, True, False, True, CStr(txtDisease.Text))
        If Not rsTemp Is Nothing Then
            txtDisease.Tag = rsTemp!ID: txtDisease.Text = rsTemp!����
        Else
            zlControl.TxtSelAll txtDisease
        End If
    End If
End Sub
Private Sub txtSoElContent_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "," Then
        KeyAscii = 0
    End If
End Sub
