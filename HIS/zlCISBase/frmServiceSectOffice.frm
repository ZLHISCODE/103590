VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmServiceSectOffice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�洢�ⷿ����"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmServiceSectOffice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDrug 
      Height          =   3300
      Left            =   2160
      ScaleHeight     =   3240
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2790
         Left            =   45
         TabIndex        =   19
         Top             =   405
         Visible         =   0   'False
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   83
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   21
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   50
         TabIndex        =   20
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdMedi 
      Caption         =   "��"
      Height          =   285
      Left            =   7080
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   548
      Width           =   285
   End
   Begin VB.Frame frame 
      Caption         =   "Ӧ����ͬԺ��(&B)"
      Height          =   1245
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   8295
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ����������ҩƷ(&6)"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   16
         Top             =   960
         Width           =   3165
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ����ͬ��������ҩƷ(&5)"
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   15
         Top             =   930
         Width           =   3255
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ�Ʒ��������ҩƷ(&2)"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   3555
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�Ƭ������ҩƷ(&4)"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   7
         Top             =   600
         Width           =   4545
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�����ҩ��(&3)"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   600
         Width           =   3345
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "��Ӧ���ڱ����ҩƷ(&1)"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   2595
      Picture         =   "frmServiceSectOffice.frx":000C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&C)"
      Height          =   350
      Left            =   1305
      Picture         =   "frmServiceSectOffice.frx":0156
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1290
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   2
      Top             =   540
      Width           =   4980
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   6210
      TabIndex        =   8
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      Picture         =   "frmServiceSectOffice.frx":02A0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7320
      TabIndex        =   9
      Top             =   5625
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2925
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":03EA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":0984
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":0F1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfServiceSectOffice 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmServiceSectOffice.frx":2C28
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "���      �����̣�       ��λ��ƿ"
      Height          =   180
      Left            =   2130
      TabIndex        =   14
      Top             =   930
      Width           =   3150
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "ָ��ҩƷ(&M)"
      Height          =   180
      Left            =   1095
      TabIndex        =   1
      Top             =   600
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ѡ����ҩƷ�����ø�ҩƷ�Ĵ洢�ⷿ�Լ�ҩ���ķ�����ҡ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   750
      TabIndex        =   0
      Top             =   240
      Width           =   5580
   End
End
Attribute VB_Name = "frmServiceSectOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln�༭ As Boolean
Private mstr���� As String
Private mlngҩƷID As Long                      'ҩƷID
Private mint���ʷ��� As Integer                 '5-����ҩ;6-�г�ҩ;7-�в�ҩ
Dim objItem As ListItem
Dim rsTemp As New ADODB.Recordset
Private mstrPrivs As String
Private mstrPara  As String         '��û��"���пⷿ"Ȩ��ʱ��¼��ǰҩƷ�����洢�ⷿ���
Private bln��ҩ��ҩ�����ʲ��� As Boolean
Private mstr�����ⷿID As String
Private mstrȫ���ⷿID As String
Private mstrStationNo As String
Private mrs���� As ADODB.Recordset
Private mstr������� As String
Private mintRow As Integer      '��¼��ǰ��
Private mintFind As Integer '������¼��ѯ���ĸ�λ����

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With lvwItems
        For i = 1 To .ListItems.Count
            If chkAllSelect.Value = 1 Then
                .ListItems(i).Checked = True
            Else
                .ListItems(i).Checked = False
            End If
        Next
    End With
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long, lngRows As Long
    With msfServiceSectOffice
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = ""
            .TextMatrix(lngRow, 3) = ""
            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMedi_Click()
    err = 0: On Error GoTo ErrHand
    Call AddColumnHeader
    
    gstrSql = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I" & _
            " where I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mint���ʷ���)
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "��δ��������������ҩƷ��", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If mint���ʷ��� <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "�����̣�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.picDrug
        lblFind.Visible = False
        txtFind.Visible = False
        chkAllSelect.Visible = False
        lvwItems.Move 50, 0, .Width - 100, .Height
        
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        lvwItems.Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub cmdSave_Click()
    Dim strPara As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    If mlngҩƷID = 0 Then
        MsgBox "����ѡ��ҩƷ��", vbInformation, gstrSysName
        txtMedi.SetFocus
        Exit Sub
    End If
    If msfServiceSectOffice.Active = False Then
        MsgBox "û���ҵ��κ�ҩƷ�ⷿ�����ڲ��Ź��������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If optӦ����(0).Value = False Then
        For i = 1 To optӦ����.UBound
            If optӦ����(i).Value = True Then
                If MsgBox("��ҩƷ���õĴ洢�ⷿӦ�÷�ΧΪ��" & optӦ����(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    '�������봮������
    lngRows = msfServiceSectOffice.Rows - 1
    For lngRow = 1 To lngRows
        If msfServiceSectOffice.TextMatrix(lngRow, 1) <> "" Then
            strPara = strPara & "!!" & msfServiceSectOffice.RowData(lngRow)
            strPara = strPara & "|" & msfServiceSectOffice.TextMatrix(lngRow, 4)
        End If
    Next
    If strPara <> "" Then
        strPara = Mid(strPara, 3)
        
        If mstrPara <> "" Then
            strPara = strPara & mstrPara
        End If
    ElseIf mstrPara <> "" Then
        strPara = Mid(mstrPara, 3)
    End If
        
    gstrSql = "zl_ҩƷ�洢�ⷿ_UPDATE(" & mlngҩƷID & ",'" & strPara & "'"
    If optӦ����(0).Value Then
        gstrSql = gstrSql & ",1)"
    ElseIf optӦ����(1).Value Then
        gstrSql = gstrSql & ",2)"
    ElseIf optӦ����(2).Value Then
        gstrSql = gstrSql & ",3)"
    ElseIf optӦ����(3).Value Then
        gstrSql = gstrSql & ",4)"
    ElseIf optӦ����(4).Value Then
        gstrSql = gstrSql & ",5)"
    Else
        gstrSql = gstrSql & ",6)"
    End If
    Call zldatabase.ExecuteProcedure(gstrSql, "����ҩƷ�洢�ⷿ�ͷ������")
    
    MsgBox "��ҩƷ�Ĵ洢�ⷿ�ͷ�����ұ���ɹ���", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If picDrug.Visible = True Then
            picDrug.Visible = False
            With msfServiceSectOffice
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strRange As String
    Dim n As Integer
    
    Call InitFace
    Call ShowData
    
    mintFind = 1
    If bln��ҩ��ҩ�����ʲ��� = True Then
        MsgBox "�������þ���ҩ��ҩ�����ʵĲ��š�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If mbln�༭ = False Then
        msfServiceSectOffice.Active = False
        frame.Enabled = False
        cmdSave.Visible = False
        cmdRestore.Visible = False
        cmdClear.Visible = False
        cmdClose.Caption = "�˳�(&X)"
    End If
    
    strRange = zldatabase.GetPara("Ӧ�÷�Χ", glngSys, 1023, False)
    For n = 0 To optӦ����.Count - 1
        optӦ����(n).Enabled = Mid(strRange, n + 1, 1)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs���� = Nothing
    mintRow = 0
End Sub

Private Sub lvwItems_GotFocus()
    Dim j As Integer
    
    With lvwItems
        For j = 1 To .ListItems.Count
            .ListItems(j).ForeColor = vbBlack
        Next
    End With
End Sub

Private Sub msfServiceSectOffice_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    MsgBox "msfServiceSectOffice_BeforeDeleteRow"
    msfServiceSectOffice.TextMatrix(Row, 1) = ""
    msfServiceSectOffice.TextMatrix(Row, 3) = ""
    msfServiceSectOffice.TextMatrix(Row, 4) = ""
    Cancel = True
End Sub

Private Sub msfServiceSectOffice_CommandClick()
'    Dim str������� As String
'    Dim objItem As ListItem
    Dim bln���� As Boolean
    
    bln���� = Check�������
    If bln���� = True Then Exit Sub
    mintRow = msfServiceSectOffice.Row
    
    Call frmServiceSelect.ShowMe(frmServiceSectOffice, mintRow, mstr�������, 1)
End Sub

Private Function Check�������() As Boolean
    '���ܣ���鵱ǰ�ⷿ�ǲ���ҩ�������Ƿ������ٴ�����
    '����ֵ true ��ǰ�ⷿ����ҩ��Ҳû�������ٴ�����,false ��ǰ�ⷿ��ҩ�����߻����������ٴ�����
    Dim str������� As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    cmdSave.Enabled = True
    str������� = ""
    gstrSql = "select distinct ������� from ��������˵�� where ����ID=[1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡ�������", msfServiceSectOffice.RowData(msfServiceSectOffice.Row))
    
    Do While Not rsTemp.EOF
        str������� = str������� & "," & rsTemp!�������
        rsTemp.MoveNext
    Loop
    If str������� <> "" Then
        str������� = Mid(str�������, 2)
        If InStr(1, str�������, 3) <> 0 Then
            str������� = "0,1,2,3"
        ElseIf InStr(1, str�������, 1) <> 0 Or InStr(1, str�������, 2) <> 0 Then
            str������� = str������� & ",3"
        End If
    Else
        str������� = "0"
    End If

    mstr������� = str�������
    gstrSql = " Select ID,����,����,���� From ���ű� A,��������˵�� B " & _
              " Where A.ID=B.����ID And B.�������� Like '�ٴ�%'" & _
              " And Instr([1], ',' || B.������� || ',') > 0"
    gstrSql = gstrSql & " and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) "
    
    If mstrStationNo <> "" Then
            gstrSql = gstrSql & " And (A.վ�� = '" & mstrStationNo & "' Or A.վ�� is Null) "
        End If
    Set mrs���� = zldatabase.OpenSQLRecord(gstrSql, "��ȡ�������", "," & str������� & ",")
    
    If mrs����.RecordCount = 0 Then
        MsgBox "��ǰ�ⷿ����ҩ������δ�����ٴ����ң�[���Ź���]", vbInformation, gstrSysName
        msfServiceSectOffice.Text = ""
        msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Row, 4) = ""
        Check������� = True
        Exit Function
    End If
    Check������� = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub msfServiceSectOffice_EnterCell(Row As Long, Col As Long)
    If Col = 3 Then
        msfServiceSectOffice.TxtEnable = True
    End If
End Sub

Private Sub msfServiceSectOffice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim bln���� As Boolean
    Dim objItem As ListItem
    Dim rsRecord As ADODB.Recordset
    Dim strKey As String
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        With msfServiceSectOffice
            If .Col = 3 Then
                strKey = Trim(UCase(.Text))

                If strKey = "" Then Exit Sub
                Debug.Print strKey
                mintRow = .Row
                
                bln���� = Check�������
                If bln���� = True Then
                    .TextMatrix(.Row, 3) = ""
                    .TextMatrix(.Row, 4) = ""
                Else
                    gstrSql = " Select distinct A.ID,A.����,A.����,A.���� From ���ű� A,��������˵�� B,�������ʷ��� C " & _
                        " Where A.ID=B.����ID And B.��������=C.���� And Instr('3ABCDEF', C.����) > 0 " & _
                        " And Instr([1], ',' || B.������� || ',') > 0"
                        
                    gstrSql = gstrSql & " and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) and ( a.���� like [2] or a.���� like [2] or a.���� like [2]) "
                    Set rsRecord = zldatabase.OpenSQLRecord(gstrSql, "����", "," & mstr������� & ",", strKey & "%")
                    
                    If rsRecord.RecordCount > 1 Then
                        mintRow = .Row
                        Call frmServiceSelect.ShowMe(frmServiceSectOffice, mintRow, mstr�������, 1, strKey)
                    ElseIf rsRecord.RecordCount = 1 Then
                        .Text = IIf(IsNull(rsRecord!����), "", rsRecord!����)
                        .TextMatrix(msfServiceSectOffice.Row, 4) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                        .TextMatrix(msfServiceSectOffice.Row, 3) = .Text
                        .SetFocus
                    ElseIf rsRecord.RecordCount = 0 Then
                        MsgBox "û���ҵ���Ӧ�Ĳ��ţ�", vbInformation, gstrSysName
                        .Text = ""
                        .TextMatrix(msfServiceSectOffice.Row, 3) = msfServiceSectOffice.Text
                        .TextMatrix(msfServiceSectOffice.Row, 4) = ""
                        .TxtSetFocus
                        Cancel = True
                    End If
                End If
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
                    
Private Sub optӦ����_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optӦ����.UBound
        If i = Index Then
            optӦ����(i).FontBold = True
        Else
            optӦ����(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim i As Integer
    Dim blnResult As Boolean
    Dim j As Integer
    Dim k As Integer
    
    blnResult = False
    With lvwItems
        If KeyCode = vbKeyReturn And Trim(txtFind.Text) <> "" Then
            strFind = UCase(Trim(txtFind.Text))
            If mintFind > .ListItems.Count Then
                mintFind = 1
            Else
                mintFind = mintFind + 1
                If mintFind > .ListItems.Count Then
                    mintFind = 1
                End If
            End If
            
            For i = mintFind To .ListItems.Count
                If IsNumeric(strFind) Then
                    If .ListItems(i).ListSubItems(.ColumnHeaders("����").Index - 1).Text = strFind Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        
                        mintFind = i
                        Exit Sub
                    End If
                    
                    If i = .ListItems.Count Then
                        For k = 1 To mintFind
                            If .ListItems(k).ListSubItems(.ColumnHeaders("����").Index - 1).Text = strFind Then
                                .ListItems(k).EnsureVisible
                                For j = 1 To .ListItems.Count
                                    .ListItems(j).ForeColor = vbBlack
                                Next
                                .ListItems(k).ForeColor = vbBlue
                                
                                mintFind = k
                                Exit Sub
                            End If
                        Next
                    End If
                Else
                    If .ListItems(i).ListSubItems(.ColumnHeaders("����").Index - 1).Text Like "*" & strFind & "*" Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        mintFind = i
                        blnResult = True
                        Exit Sub
                    End If
                    
                    If i = .ListItems.Count Then
                        For k = 1 To mintFind
                            If .ListItems(k).ListSubItems(.ColumnHeaders("����").Index - 1).Text Like "*" & strFind & "*" Then
                                .ListItems(k).EnsureVisible
                                For j = 1 To .ListItems.Count
                                    .ListItems(j).ForeColor = vbBlack
                                Next
                                .ListItems(k).ForeColor = vbBlue
                                
                                mintFind = k
                                blnResult = True
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            Next
            
            If blnResult = False Then
                For i = mintFind To .ListItems.Count
                    If .ListItems(i).Text Like "*" & strFind & "*" Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        mintFind = i
                        blnResult = True
                        Exit Sub
                    End If
                Next
                
                For k = 1 To mintFind
                    If .ListItems(k).Text Like "*" & strFind & "*" Then
                        .ListItems(k).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(k).ForeColor = vbBlue
                        
                        mintFind = k
                        blnResult = True
                        Exit Sub
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
            " where I.ID=N.�շ�ϸĿID and I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mint���ʷ���, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "δ�ҵ�ָ������ҩƷ��������ָ����", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If mint���ʷ��� <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "�����̣�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Call AddColumnHeader
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDrug
        lblFind.Visible = False
        txtFind.Visible = False
        chkAllSelect.Visible = False
        lvwItems.Move 50, 0, .Width - 100, .Height
        
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        lvwItems.Visible = True
        .SetFocus
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub ShowData()
    '��ȡ���ݲ���ʾ����
    Const str��ҩ As String = "'��ҩ%'"
    Const str��ҩ As String = "'��ҩ%'"
    Const str��ҩ As String = "'��ҩ%'"
    Dim str�ⷿID As String, str���� As String, str����ID As String
    Dim intRow As Integer, intRows As Integer
    Dim blnSel As Boolean
    Dim lng������ĿID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsOther As New ADODB.Recordset
    Dim dbl���пⷿ As Boolean
    Dim strTmp As String
    Dim strIdArr
    
    err = 0: On Error GoTo ErrHand
    
    mstrPara = ""
    If InStr(1, ";" & mstrPrivs & ";", ";���пⷿ;") > 0 Then dbl���пⷿ = True
    
    Call cmdClear_Click
    If mblnFirst Then
        '��ȡҩƷ��Ϣ
        If mlngҩƷID <> 0 Then
            gstrSql = " Select A.ҩ��ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ,B.ҩƷ���� ���� " & _
                      " From ҩƷ��� A,ҩƷ���� B,�շ���ĿĿ¼ I " & _
                      " Where A.ҩ��ID=B.ҩ��ID And A.ҩƷID=I.ID And A.ҩƷID=[1] "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡҩƷ��Ϣ", mlngҩƷID)
            
            lng������ĿID = rsTemp!ҩ��ID
            txtMedi.Tag = "[" & rsTemp!���� & "]" & rsTemp!����
            txtMedi.Text = txtMedi.Tag
            mstr���� = rsTemp!����
            If mint���ʷ��� <> "7" Then
                Me.lblSpec.Caption = "���" & IIf(IsNull(rsTemp!���), "", rsTemp!���) & _
                    "   �����̣�" & IIf(IsNull(rsTemp!����), "", rsTemp!����) & _
                    "   ��λ��" & IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
            Else
                Me.lblSpec.Caption = "�����̣�" & IIf(IsNull(rsTemp!����), "", rsTemp!����) & "   ��λ��" & IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
            End If
        End If
        
        '����ҩƷ����;������ȡ������洢�Ŀⷿ
        gstrSql = " Select ID,����,���� From ���ű� " & _
                  " Where  (����ʱ�� is null or to_char(����ʱ��,'yyyy-mm-dd')='3000-01-01') and ID in (select distinct ����id from ��������˵�� where �������� like "
        If mint���ʷ��� = "5" Then
            gstrSql = gstrSql & str��ҩ
        ElseIf mint���ʷ��� = "6" Then
            gstrSql = gstrSql & str��ҩ
        Else
            gstrSql = gstrSql & str��ҩ
        End If
        gstrSql = gstrSql & " or ��������='�Ƽ���')"
        
        gstrSql = gstrSql & " and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) "
        
        If mstrStationNo <> "" Then
            gstrSql = gstrSql & " And (վ�� = '" & mstrStationNo & "' Or վ�� is Null) "
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "����ҩƷ����;������ȡ������洢�Ŀⷿ(�����ⷿ)")
        mstrȫ���ⷿID = ""
        Do While Not rsTemp.EOF
            mstrȫ���ⷿID = mstrȫ���ⷿID & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
        If mstrȫ���ⷿID <> "" Then
            mstrȫ���ⷿID = Mid(mstrȫ���ⷿID, 2)
            bln��ҩ��ҩ�����ʲ��� = False
        Else
            bln��ҩ��ҩ�����ʲ��� = True
            Exit Sub
        End If
        
        strTmp = gstrSql
        If Not dbl���пⷿ Then
            '��ȡ�����ⷿ
            gstrSql = strTmp & " And Id not In(Select ����ID From ������Ա Where ��Աid=[1])"
            Set rsOther = zldatabase.OpenSQLRecord(gstrSql, "����ҩƷ����;������ȡ������洢�Ŀⷿ(�����ⷿ)", UserInfo.ID)
            
            mstr�����ⷿID = ""
            Do While Not rsOther.EOF
                mstr�����ⷿID = mstr�����ⷿID & "," & rsOther!ID
                rsOther.MoveNext
            Loop
            If mstr�����ⷿID <> "" Then
                mstr�����ⷿID = Mid(mstr�����ⷿID, 2)
            End If
                        
            'ȡ��ǰ�û������ⷿ
            gstrSql = strTmp & " And Id In(Select ����ID From ������Ա Where ��Աid=[1])"
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "����ҩƷ����;������ȡ������洢�Ŀⷿ", UserInfo.ID)
                
        Do While Not rsTemp.EOF
            msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Rows - 1, 2) = rsTemp!����
            msfServiceSectOffice.RowData(msfServiceSectOffice.Rows - 1) = rsTemp!ID
            msfServiceSectOffice.Rows = msfServiceSectOffice.Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
        If str�ⷿID <> "" Then
            str�ⷿID = Mid(str�ⷿID, 2)
            msfServiceSectOffice.Rows = msfServiceSectOffice.Rows - 1
            msfServiceSectOffice.Active = True
        Else
            msfServiceSectOffice.Active = False
        End If
    End If
    
    'ȡ���пⷿ
    str�ⷿID = ""
    intRows = msfServiceSectOffice.Rows - 1
    For intRow = 1 To intRows
        str�ⷿID = str�ⷿID & "," & msfServiceSectOffice.RowData(intRow)
    Next
    If str�ⷿID <> "" Then str�ⷿID = Mid(str�ⷿID, 2)
    
   '����Ӧ������֯��װ�뵥�ݿؼ�
    gstrSql = " Select A.�շ�ϸĿID,A.��������ID,A.ִ�п���ID,B.���� From �շ�ִ�п��� A,���ű� B " & _
              " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=[1] And Instr([2],','||A.ִ�п���ID||',') > 0 " & _
              " Order by A.ִ�п���ID"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡ�����õ��շ�ִ�п�������", mlngҩƷID, "," & mstrȫ���ⷿID & ",")
    
    If rsTemp.RecordCount = 0 And mblnFirst And mlngҩƷID <> 0 Then
        '��ȡͬƷ�����������ҩƷ��������Ϊȱʡ����
        gstrSql = " Select A.��������ID,A.ִ�п���ID,B.���� From �շ�ִ�п��� A,���ű� B," & _
                  "     (Select ҩƷID From ҩƷ��� Where ҩ��ID=[1] And Rownum<2) C" & _
                  " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ҩƷID And Instr([2],','||A.ִ�п���ID||',') > 0 " & _
                  " Order by A.ִ�п���ID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡ��ͬ���ʷ�������ͬ���͵�ҩƷ��������Ϊȱʡ����", lng������ĿID, "," & mstrȫ���ⷿID & ",")
                
        If rsTemp.RecordCount = 0 Then
            '��ȡ��ͬ���ʷ�������ͬ���͵�ҩƷ��������Ϊȱʡ����
            gstrSql = " Select A.��������ID,A.ִ�п���ID,B.���� From �շ�ִ�п��� A,���ű� B," & _
                      "     (Select C.ҩƷID From ������ĿĿ¼ A,ҩƷ���� B,ҩƷ��� C,�շ�ִ�п��� D " & _
                      "     Where A.ID=B.ҩ��ID And B.ҩ��ID=C.ҩ��ID And C.ҩƷID=D.�շ�ϸĿID And B.ҩƷ����=[1] And A.���=[2] And Rownum<2) C" & _
                      " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ҩƷID And Instr([3],','||A.ִ�п���ID||',') > 0 " & _
                      " Order by A.ִ�п���ID"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡ��ͬ���ʷ�������ͬ���͵�ҩƷ��������Ϊȱʡ����", mstr����, mint���ʷ���, "," & mstrȫ���ⷿID & ",")
            
            If rsTemp.RecordCount <> 0 Then
                MsgBox "��ǰ���ҩƷδ���ô洢�ⷿ����ȡ������ͬ�Ĺ��ҩƷ�Ĵ洢�ⷿ��Ϊȱʡ���ݣ�", vbInformation, gstrSysName
            End If
        Else
            MsgBox "��ǰ���ҩƷδ���ô洢�ⷿ����ȡͬƷ���¹��ҩƷ�Ĵ洢�ⷿ��Ϊȱʡ���ݣ�", vbInformation, gstrSysName
        End If
    End If
    For intRow = 1 To intRows
        str���� = "": str����ID = ""
        rsTemp.Filter = "ִ�п���ID=" & msfServiceSectOffice.RowData(intRow)
        
        blnSel = False
        Do While Not rsTemp.EOF
            blnSel = True
            str���� = str���� & "," & Nvl(rsTemp!����)
            str����ID = str����ID & "," & Nvl(rsTemp!��������ID, 0)
            rsTemp.MoveNext
        Loop
        If str���� <> "" Then
            str���� = Mid(str����, 2)
            str����ID = Mid(str����ID, 2)
            If str����ID = "0" Then str����ID = ""
        End If
        msfServiceSectOffice.TextMatrix(intRow, 0) = intRow
        If blnSel Then msfServiceSectOffice.TextMatrix(intRow, 1) = "��"
        msfServiceSectOffice.TextMatrix(intRow, 3) = str����
        msfServiceSectOffice.TextMatrix(intRow, 4) = str����ID
    Next
    
    'ȡ����ִ�п���
    If Not dbl���пⷿ And mstr�����ⷿID <> "" Then
        gstrSql = " Select DISTINCT ��������ID,ִ�п���ID From �շ�ִ�п��� " & _
              " Where �շ�ϸĿID=[1] And Instr([2],','||ִ�п���ID||',') > 0 " & _
              " Order by ִ�п���ID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��ȡ�����õ��շ�ִ�п�������", mlngҩƷID, "," & mstr�����ⷿID & ",")
                        
        strIdArr = Split(mstr�����ⷿID, ",")
        For intRow = 0 To UBound(strIdArr)
            str����ID = ""
            rsTemp.Filter = "ִ�п���ID=" & strIdArr(intRow)
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str����ID = str����ID & "," & Nvl(rsTemp!��������ID, 0)
                rsTemp.MoveNext
            Loop
            If str����ID <> "" Then
                str����ID = Mid(str����ID, 2)
                If str����ID = "0" Then str����ID = ""
                mstrPara = mstrPara & "!!" & CStr(strIdArr(intRow))
                mstrPara = mstrPara & "|" & str����ID
            End If
        Next
    End If
    
    '�޸�Ӧ������Ϣ
    optӦ����(2).Caption = "Ӧ�������С�" & Switch(mint���ʷ��� = 5, "����ҩ", mint���ʷ��� = 6, "�г�ҩ", mint���ʷ��� = 7, "�в�ҩ") & "��(&3)"
    optӦ����(3).Caption = "Ӧ�������С�" & mstr���� & "��������ҩƷ(&4)"
    
    If mint���ʷ��� = 7 Then optӦ����(3).Enabled = False
    mblnFirst = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '��ʼ���ؼ�
    With msfServiceSectOffice
        .Rows = 2
        .Cols = 5
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "ѡ��"
        .TextMatrix(0, 2) = "�洢�ⷿ"
        .TextMatrix(0, 3) = "�������"
        .TextMatrix(0, 4) = "�������ID"
        .TextMatrix(1, 0) = "1"
        .colData(0) = 5
        .colData(1) = -1
        .colData(2) = 5
        .colData(3) = 1
        .colData(4) = 5
        .ColWidth(0) = 300
        .ColWidth(1) = 500
        .ColWidth(2) = 1500
        .ColWidth(3) = 5000
        .ColWidth(4) = 0
        
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub msfServiceSectOffice_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    MsgBox "msfServiceSectOffice_AfterAddRow"
    '�޸������
    MsgBox "msfServiceSectOffice_AfterAddRow"
    With msfServiceSectOffice
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfServiceSectOffice_AfterDeleteRow()
    Dim lngCurRow As Long
    
    MsgBox "msfServiceSectOffice_AfterDeleteRow"
    '�޸������
    With msfServiceSectOffice
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub AddColumnHeader(Optional ByVal blnҩƷ As Boolean = True)
    If blnҩƷ Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 2000
            .Add , "����", "����", 1000
            .Add , "���", "���", 800
            .Add , "����", "������", 1500
            .Add , "��λ", "��λ", 800
        End With
        With Me.lvwItems
            .Checkboxes = False
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "1"
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 2000
            .Add , "����", "����", 1000
            .Add , "����", "����", 1000
        End With
        With Me.lvwItems
            .Checkboxes = True
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 2
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "2"
    End If
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str���� As String, str����ID As String
    
    If lvwItems.Tag = "1" Then
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        With Me.lvwItems
            If mlngҩƷID <> Mid(.SelectedItem.Key, 2) Then
                mlngҩƷID = Mid(.SelectedItem.Key, 2)
                Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                Me.txtMedi.Text = Me.txtMedi.Tag
                mstr���� = .SelectedItem.SubItems(3)
                If mint���ʷ��� <> "7" Then
                    Me.lblSpec.Caption = "���" & .SelectedItem.SubItems(2) & _
                        "   �����̣�" & .SelectedItem.SubItems(3) & _
                        "   ��λ��" & .SelectedItem.SubItems(4)
                Else
                    Me.lblSpec.Caption = "�����̣�" & .SelectedItem.SubItems(3) & "   ��λ��" & .SelectedItem.SubItems(4)
                End If
                Call ShowData
            End If
            Me.txtMedi.SetFocus
            Call OS.PressKey(vbKeyTab)
        End With
        picDrug.Visible = False
    Else
        'ѭ����ȡ�û���ѡ��Ŀ���
        lngRows = lvwItems.ListItems.Count
        For lngRow = 1 To lngRows
            If lvwItems.ListItems(lngRow).Checked Then
                str���� = str���� & "," & lvwItems.ListItems(lngRow).Text
                str����ID = str����ID & "," & Mid(lvwItems.ListItems(lngRow).Key, 2)
            End If
        Next
        If str���� <> "" Then
            str���� = Mid(str����, 2)
            str����ID = Mid(str����ID, 2)
        End If
        msfServiceSectOffice.Visible = True
        lvwItems.Visible = False
        picDrug.Visible = False
        If str���� <> "" Then msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Row, 1) = "��"
        msfServiceSectOffice.Text = str����
        msfServiceSectOffice.TextMatrix(mintRow, 3) = msfServiceSectOffice.Text
        msfServiceSectOffice.TextMatrix(mintRow, 4) = str����ID
        If msfServiceSectOffice.Rows - 1 > msfServiceSectOffice.Row Then msfServiceSectOffice.Row = msfServiceSectOffice.Row + 1: msfServiceSectOffice.SetFocus
    End If
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

'Private Sub lvwItems_LostFocus()
'    Me.lvwItems.Visible = False
'End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal lngҩƷID As Long, ByVal int��;���� As Integer, ByVal bln�༭ As Boolean, ByVal strStationNo As String)
    On Error Resume Next
    mblnFirst = True
    mlngҩƷID = lngҩƷID
    mint���ʷ��� = int��;����
    mbln�༭ = bln�༭
    'mstrPrivs = gstrPrivs
    mstrPrivs = frmParent.mstrPrivs
    mstrStationNo = strStationNo
    Me.Show 1, frmParent
End Sub
